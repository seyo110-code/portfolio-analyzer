"""
AI 모델 설정 관리
- Provider 전환: Gemini ↔ GitHub Models
- Fast / Pro 모델 개별 선택 → .env 자동 업데이트
- GitHub 전환 시 gh auth token으로 토큰 자동 주입

사용법: python3 ai_model.py
"""
import re
import subprocess
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

ENV_FILE = Path(".env")

GEMINI_MODELS = [
    # ── Gemini 3 (최신 preview) ───────────────────────────────
    ("gemini-3.1-pro-preview", "3.1 Pro preview   — 최신 최고 추론, 에이전틱"),
    ("gemini-3-flash-preview", "3.0 Flash preview — 프론티어급 성능, 빠름"),
    # ── Gemini 2.5 (안정) ─────────────────────────────────────
    ("gemini-2.5-pro",         "2.5 Pro           — 안정화 최고 정확도  (분석 권장)"),
    ("gemini-2.5-flash",       "2.5 Flash         — 가성비 최고         (기본값)"),
    ("gemini-2.5-flash-lite",  "2.5 Flash Lite    — 가장 빠름, 저비용"),
    # ── 퇴역 예정 ─────────────────────────────────────────────
    ("gemini-2.0-flash",       "2.0 Flash         — ⚠️  2026-06 퇴역 예정"),
]

GITHUB_MODELS = [
    # ── OpenAI GPT-5 ──────────────────────────────────────────
    ("openai/gpt-5",           "GPT-5            — OpenAI 최신 플래그십"),
    ("openai/gpt-5-mini",      "GPT-5 mini       — GPT-5 경량, 빠름 (권장)"),
    # ── OpenAI GPT-4 ──────────────────────────────────────────
    ("openai/gpt-4.1",         "GPT-4.1          — 안정화 플래그십"),
    ("openai/gpt-4o",          "GPT-4o           — 멀티모달, 고성능"),
    ("openai/gpt-4o-mini",     "GPT-4o mini      — 빠름, 저비용"),
    ("openai/o3-mini",         "o3-mini          — 추론 특화"),
    # ── Meta ──────────────────────────────────────────────────
    ("meta/llama-3.3-70b-instruct", "Llama 3.3 70B — Meta 오픈소스 최신"),
    # ── DeepSeek ──────────────────────────────────────────────
    ("deepseek/deepseek-r1",   "DeepSeek R1      — 추론 특화 오픈소스"),
]

PROVIDER_DEFAULTS = {
    "gemini": ("gemini-2.5-flash", "gemini-2.5-pro"),
    "github": ("openai/gpt-5-mini", "openai/gpt-5-mini"),
}

TEST_TIMEOUT = 10  # 초


# ─── .env 읽기/쓰기 ──────────────────────────────────────────

def read_env() -> dict[str, str]:
    env = {}
    if ENV_FILE.exists():
        for line in ENV_FILE.read_text().splitlines():
            m = re.match(r'^([A-Z_]+)\s*=\s*(.*)$', line.strip())
            if m:
                env[m.group(1)] = m.group(2)
    return env


def update_env(key: str, value: str):
    lines = ENV_FILE.read_text().splitlines()
    new_lines = []
    replaced = False
    for line in lines:
        # 활성/주석 모두 이 키면 처음 한 번만 새 값으로 교체, 나머지는 삭제
        if re.match(rf'^#?\s*{re.escape(key)}\s*=', line.strip()):
            if not replaced:
                new_lines.append(f'{key}={value}')
                replaced = True
        else:
            new_lines.append(line)
    if not replaced:
        new_lines.append(f'{key}={value}')
    ENV_FILE.write_text('\n'.join(new_lines) + '\n')


# ─── GitHub 인증 처리 ─────────────────────────────────────────

def check_github_auth() -> tuple[bool, str]:
    """(로그인 여부, 토큰 종류) 반환
    토큰 종류: 'classic' | 'fine-grained' | 'none'
    - classic (gho_...): models:read 스코프 없어도 GitHub Models 사용 가능
    - fine-grained (github_pat_...): models:read 스코프 필요
    """
    try:
        result = subprocess.run(
            ["gh", "auth", "status"], capture_output=True, text=True
        )
        output = result.stdout + result.stderr
        logged_in = "Logged in" in output

        if not logged_in:
            return False, "none"

        # 토큰 종류 파악
        if "gho_" in output:
            token_type = "classic"
        elif "github_pat_" in output:
            has_models = "models:read" in output
            token_type = "fine-grained-ok" if has_models else "fine-grained-noscope"
        else:
            token_type = "classic"  # 기본값: classic으로 가정
        return True, token_type
    except FileNotFoundError:
        return False, "none"


def get_github_token() -> str | None:
    try:
        result = subprocess.run(
            ["gh", "auth", "token"], capture_output=True, text=True
        )
        token = result.stdout.strip()
        return token if token else None
    except FileNotFoundError:
        return None


def setup_github_provider() -> bool:
    """GitHub 전환 준비. 성공하면 True 반환."""
    print("\n  🔍 GitHub 인증 상태 확인 중...")
    logged_in, token_type = check_github_auth()

    if not logged_in:
        print("  ❌ GitHub 로그인이 필요합니다.")
        print("     → 아래 명령 실행 후 다시 시도하세요:")
        print("       gh auth login")
        return False

    if token_type == "fine-grained-noscope":
        # Fine-grained PAT인데 models:read 없는 경우만 경고
        print("  ⚠️  Fine-grained PAT에 models:read 권한이 없습니다.")
        print("     → GitHub Settings > Personal access tokens 에서")
        print("       models:read 권한을 추가하거나, classic 토큰으로 전환하세요.")
        print("     → classic 토큰은 models:read 없이도 사용 가능합니다.")
        choice = input("\n  그래도 계속 진행할까요? (y/n): ").strip().lower()
        if choice != "y":
            return False
    elif token_type == "classic":
        print("  ✅ Classic OAuth 토큰 — models:read 스코프 없이도 GitHub Models 사용 가능")
    elif token_type == "fine-grained-ok":
        print("  ✅ Fine-grained PAT (models:read 권한 확인됨)")

    token = get_github_token()
    if not token:
        print("  ❌ 토큰을 가져올 수 없습니다.")
        return False

    update_env("GITHUB_TOKEN", token)
    print(f"  ✅ GitHub 토큰 자동 주입 완료 ({token[:8]}...)")

    # 연결 빠른 확인 (gh CLI로 직접 테스트)
    print("  🔗 GitHub Models 연결 확인 중...", end=" ", flush=True)
    try:
        r = subprocess.run(
            ["curl", "-s", "-o", "/dev/null", "-w", "%{http_code}",
             "-X", "POST",
             "-H", f"Authorization: Bearer {token}",
             "-H", "Content-Type: application/json",
             "https://api.githubcopilot.com/chat/completions",
             "-H", "Copilot-Integration-Id: vscode-chat",
             "-d", '{"model":"openai/gpt-4o-mini","messages":[{"role":"user","content":"hi"}]}'],
            capture_output=True, text=True, timeout=10
        )
        code = r.stdout.strip()
        if code == "200":
            print("✅ 연결 OK")
        elif code == "403":
            print("⚠️  연결됨 (403 — 무료 할당량 초과 또는 권한 부족)")
            print("     → https://github.com/settings/billing/summary 에서 확인")
        elif code == "401":
            print("❌ 인증 실패 (401)")
        else:
            print(f"⚠️  HTTP {code}")
    except Exception:
        print("(확인 실패 — 네트워크 오류)")

    return True


# ─── 모델 테스트 ─────────────────────────────────────────────

def test_model(model: str) -> tuple[bool, str]:
    """모델이 실제로 응답하는지 확인. (성공 여부, 메시지) 반환"""
    import signal

    def _timeout_handler(signum, frame):
        raise TimeoutError()

    print(f"  🧪 테스트 중: {model} ...", end=" ", flush=True)
    try:
        from ai_client import get_client
        client = get_client()

        signal.signal(signal.SIGALRM, _timeout_handler)
        signal.alarm(TEST_TIMEOUT)
        try:
            resp = client.generate("say hi in one word", model=model)
            signal.alarm(0)
        finally:
            signal.alarm(0)

        print(f"✅  ({resp.strip()[:30]})")
        return True, resp.strip()
    except TimeoutError:
        print(f"❌  ({TEST_TIMEOUT}초 타임아웃 — 응답 없음)")
        return False, "timeout"
    except Exception as e:
        msg = str(e)
        # 에러 종류별 안내
        if "budget" in msg.lower() or "budget limit" in msg.lower():
            print(f"❌  (GitHub Models 무료 할당량 초과)")
            print(f"     → https://github.com/settings/billing/summary 에서 확인")
            return False, "budget_limit"
        elif "401" in msg or "Unauthorized" in msg or "unauthorized" in msg:
            print(f"❌  (인증 실패 — 토큰 확인 필요)")
            return False, "auth_error"
        elif "403" in msg or "Forbidden" in msg:
            print(f"❌  (403 Forbidden — 권한 부족)")
            return False, "forbidden"
        elif "404" in msg or "not found" in msg.lower():
            print(f"❌  (모델을 찾을 수 없음: {model})")
            return False, "not_found"
        else:
            print(f"❌  ({msg[:60]})")
            return False, msg


# ─── 모델 선택 ───────────────────────────────────────────────

def pick_model(label: str, current: str, models: list[tuple]) -> str | None:
    print(f"\n  [{label} 모델]  현재: {current}")
    print("  " + "-" * 58)
    for i, (model, desc) in enumerate(models, start=1):
        marker = " ◀ 현재" if model == current else ""
        print(f"  {i}. {desc}{marker}")
    print("  d. 직접 입력")
    print("  " + "-" * 58)
    choice = input("  선택 (Enter = 유지): ").strip().lower()

    if choice == "":
        return None
    if choice.isdigit() and 1 <= int(choice) <= len(models):
        selected = models[int(choice) - 1][0]
        return None if selected == current else selected
    if choice == "d":
        custom = input("  모델명 직접 입력: ").strip()
        return custom if custom and custom != current else None
    print("  → 잘못된 입력, 유지")
    return None


# ─── 메인 ────────────────────────────────────────────────────

def main():
    env      = read_env()
    provider = env.get("AI_PROVIDER", "gemini").lower()
    fast_cur = env.get("AI_MODEL_FAST", PROVIDER_DEFAULTS[provider][0])
    pro_cur  = env.get("AI_MODEL_PRO",  PROVIDER_DEFAULTS[provider][1])

    print("\n" + "=" * 62)
    print("⚙️  AI 설정")
    print("=" * 62)
    print(f"  Provider  : {provider}")
    print(f"  Fast 모델  : {fast_cur}  (이미지 분석, ETF 분류)")
    print(f"  Pro  모델  : {pro_cur}  (포트폴리오 분석, 리밸런싱)")
    print("=" * 62)

    changed = False

    # ── 1. Provider 전환 여부 ──────────────────────────────────
    other = "github" if provider == "gemini" else "gemini"
    print(f"\n  현재 Provider: {provider}")
    switch = input(f"  {other}(으)로 전환할까요? (y/n, Enter = 유지): ").strip().lower()

    if switch == "y":
        if other == "github":
            if not setup_github_provider():
                print("  → Provider 전환 취소, 기존 설정 유지")
            else:
                update_env("AI_PROVIDER", "github")
                # GitHub 기본 모델로 초기화
                fast_cur, pro_cur = PROVIDER_DEFAULTS["github"]
                update_env("AI_MODEL_FAST", fast_cur)
                update_env("AI_MODEL_PRO",  pro_cur)
                provider = "github"
                print(f"  ✅ Provider → github")
                changed = True
        else:
            update_env("AI_PROVIDER", "gemini")
            fast_cur, pro_cur = PROVIDER_DEFAULTS["gemini"]
            update_env("AI_MODEL_FAST", fast_cur)
            update_env("AI_MODEL_PRO",  pro_cur)
            provider = "gemini"
            print(f"  ✅ Provider → gemini")
            changed = True

    # ── 2. 모델 선택 ──────────────────────────────────────────
    models = GITHUB_MODELS if provider == "github" else GEMINI_MODELS

    new_fast = pick_model("Fast", fast_cur, models)
    if new_fast:
        ok, _ = test_model(new_fast)
        if ok:
            update_env("AI_MODEL_FAST", new_fast)
            print(f"  ✅ Fast: {fast_cur} → {new_fast}")
            changed = True
        else:
            keep = input(f"  응답 실패. 그래도 저장할까요? (y/n): ").strip().lower()
            if keep == "y":
                update_env("AI_MODEL_FAST", new_fast)
                print(f"  ⚠️  Fast: {fast_cur} → {new_fast} (미검증)")
                changed = True
            else:
                print(f"  → Fast 모델 변경 취소")
    else:
        test_model(fast_cur)

    new_pro = pick_model("Pro", pro_cur, models)
    if new_pro:
        ok, _ = test_model(new_pro)
        if ok:
            update_env("AI_MODEL_PRO", new_pro)
            print(f"  ✅ Pro : {pro_cur} → {new_pro}")
            changed = True
        else:
            keep = input(f"  응답 실패. 그래도 저장할까요? (y/n): ").strip().lower()
            if keep == "y":
                update_env("AI_MODEL_PRO", new_pro)
                print(f"  ⚠️  Pro : {pro_cur} → {new_pro} (미검증)")
                changed = True
            else:
                print(f"  → Pro 모델 변경 취소")
    else:
        test_model(pro_cur)

    # ── 최종 요약 ─────────────────────────────────────────────
    print("\n" + "=" * 62)
    if changed:
        env = read_env()
        print(f"  Provider  : {env.get('AI_PROVIDER')}")
        print(f"  Fast 모델  : {env.get('AI_MODEL_FAST')}")
        print(f"  Pro  모델  : {env.get('AI_MODEL_PRO')}")
        print("  💾 .env 저장 완료 — 다음 실행부터 즉시 적용")
    else:
        print("  변경 없음.")
    print("=" * 62)


if __name__ == "__main__":
    main()
