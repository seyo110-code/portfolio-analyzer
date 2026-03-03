"""
AI 클라이언트 추상화 레이어
- .env의 AI_PROVIDER 설정에 따라 Gemini 또는 GitHub Models(Copilot) 사용
- 모든 스크립트는 이 모듈을 통해 AI 호출

지원 Provider:
  gemini : Google Gemini API (GEMINI_API_KEY 필요)
  github : GitHub Models / Copilot (GITHUB_TOKEN 필요)
"""
import os
import base64
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# .env 설정
AI_PROVIDER   = os.getenv("AI_PROVIDER",    "gemini")
MODEL_FAST    = os.getenv("AI_MODEL_FAST",  "gemini-2.5-flash")   # 이미지 분석, 분류
MODEL_PRO     = os.getenv("AI_MODEL_PRO",   "gemini-2.5-pro")     # 리밸런싱 분석


class AIClient:
    """공통 AI 클라이언트 인터페이스"""

    def generate(self, prompt: str, image_path: str = None, model: str = None) -> str:
        """
        텍스트 또는 이미지+텍스트 생성 요청
        :param prompt: 입력 프롬프트
        :param image_path: 이미지 파일 경로 (vision 사용 시)
        :param model: 사용할 모델 (None이면 기본값 사용)
        :return: AI 응답 텍스트
        """
        raise NotImplementedError


class GeminiClient(AIClient):
    def __init__(self):
        from google import genai
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise ValueError(".env에 GEMINI_API_KEY가 없습니다.")
        self._client = genai.Client(api_key=api_key)

    def generate(self, prompt: str, image_path: str = None, model: str = None) -> str:
        from PIL import Image
        model = model or MODEL_FAST
        contents = [prompt]
        if image_path:
            contents.append(Image.open(image_path))
        response = self._client.models.generate_content(model=model, contents=contents)
        return response.text


class GitHubClient(AIClient):
    """
    GitHub Models (https://models.github.ai/inference)
    - OpenAI 호환 API
    - GITHUB_TOKEN 필요 (gh auth token)
    - 모델명 형식: openai/gpt-5-mini, meta/llama-3.3-70b-instruct 등
    """
    ENDPOINT = "https://models.github.ai/inference"

    def __init__(self):
        from openai import OpenAI
        token = os.getenv("GITHUB_TOKEN")
        if not token:
            raise ValueError(".env에 GITHUB_TOKEN이 없습니다.")
        self._client = OpenAI(base_url=self.ENDPOINT, api_key=token)

    def generate(self, prompt: str, image_path: str = None, model: str = None) -> str:
        model = model or MODEL_FAST
        if image_path:
            # 이미지는 base64로 인코딩해서 전달
            ext  = Path(image_path).suffix.lower().replace(".", "")
            mime = f"image/{'jpeg' if ext == 'jpg' else ext}"
            with open(image_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("utf-8")
            messages = [{
                "role": "user",
                "content": [
                    {"type": "text",      "text": prompt},
                    {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                ]
            }]
        else:
            messages = [{"role": "user", "content": prompt}]

        response = self._client.chat.completions.create(model=model, messages=messages)
        return response.choices[0].message.content


def get_client() -> AIClient:
    """AI_PROVIDER 설정에 따라 적절한 클라이언트 반환"""
    provider = AI_PROVIDER.lower()
    if provider == "gemini":
        return GeminiClient()
    elif provider == "github":
        return GitHubClient()
    else:
        raise ValueError(f"지원하지 않는 AI_PROVIDER: '{provider}' (gemini / github 중 선택)")


def get_fast_model() -> str:
    return MODEL_FAST


def get_pro_model() -> str:
    return MODEL_PRO
