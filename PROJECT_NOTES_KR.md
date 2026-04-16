# NoteFlow AX PoC 진행 메모 (로컬 + 원노트)

이 문서는 지금까지 대화에서 나온 **의사결정 / 시행착오 / 현재 PoC 상태 / 다음 단계**를 한 번에 정리한 메모다.

## 목표(AX 과제 관점)
- **비정형 텍스트(메신저/메모/회의 단편)** → **LLM으로 요약/분류/액션아이템 구조화** → (가능하면) **원노트에 정리본 저장**
- 보안 관점: LLM 추론은 **로컬(Ollama)**로 처리하여 외부 LLM API로 본문을 보내지 않는 방향

## 현재까지 완료(로컬 PoC)
- **Ollama 설치 및 모델 준비**: `llama3` 사용
- **로컬에서 추론 테스트 성공**: `ollama run llama3 ...` 정상 응답
- **Python 연동**: `ollama` 파이썬 클라이언트 사용
- **프로젝트 가상환경 구성**: `.venv`에 의존성 설치
- **구조화 파이프라인**:
  - `sample.txt`(샘플 입력) → `process.py` 실행 → `out.json` 생성
  - `render_md.py` 실행 → `out.md` 생성 (원노트에 복붙 가능한 형태)

## 왜 “완전 로컬로 원노트 탭 자동 생성”이 어려웠나
- 원노트에서 **섹션(탭) 생성/이동/페이지 생성**을 자동화하려면 보통 **Microsoft Graph** 같은 **공식 API + 로그인/권한**이 필요하다.
- “계정 없이 로컬만 사용하는 원노트”는 Graph가 접근할 대상이 없다.
- 따라서 현실적인 무료/로컬 PoC는:
  - **로컬에서 구조화(out.md/out.json)** 까지 완성하고
  - 원노트에는 **수동 붙여넣기**로 데모(또는 Graph가 가능해지면 자동화 확장)

## Graph(원노트 자동 생성) 관련 이슈 요약
- 개인 계정으로 Azure Portal 로그인은 되지만,
  - **디렉터리(테넌트) 밖에서 앱 등록이 중단**되었다는 메시지로 `App registration`이 막힘.
- M365 개발자 프로그램 샌드박스(E5)는 계정 조건 때문에 “자격 없음”으로 막힐 수 있음.
- 대안:
  - **Azure 평가판(Free) 구독**을 활성화해서 **테넌트 생성 → 앱 등록 → Graph 권한**으로 진행
  - 단, 평가판은 결제수단 등록이 필요할 수 있고, **VM/DB 같은 리소스만 만들지 않으면 비용을 거의 0으로 유지**할 수 있음.

## 로컬에서 지금 당장 데모하는 방법(추천)
1) 입력을 `sample.txt`에 붙여넣거나 파일로 저장
2) 터미널에서 실행:

```bash
cd "/Users/dain/Documents/Projects/NoteFlow AX"
source .venv/bin/activate
python process.py
python render_md.py
```

3) 생성된 `out.md` 내용을 **원노트(웹/앱) 새 페이지**에 복붙

## 레포 파일 설명
- `sample.txt`: 메신저/메모가 섞인 샘플 입력
- `process.py`: 텍스트 → Ollama → JSON 구조화(`out.json`)
  - 모델이 JSON 앞뒤로 문장을 붙이는 경우를 대비해 `{...}`만 추출해 파싱하는 보조 로직 포함
  - `temperature=0`으로 결과 흔들림을 줄임
- `render_md.py`: `out.json` → 보기 좋은 노트 템플릿(`out.md`)
- `onenote_publish.py`: (Graph 가능할 때) `out.json`을 원노트 섹션에 새 페이지로 생성하는 스크립트
  - **전제**: 앱 등록(Client ID) + 권한(User.Read, Notes.ReadWrite) + 원노트가 계정 동기화 상태
- `requirements.txt`: 파이썬 의존성
- `.gitignore`: `.venv/`, `out.json`, `.env` 등 커밋 제외

## 다음 단계 옵션
### A. 로컬 PoC 완성도 올리기(권장, 빠름)
- `inputs/` 폴더에 여러 `.txt`를 넣고 **배치 처리**(`outputs/`에 `*.json`, `*.md` 생성)
- 분류 카테고리 후보를 고정해 **일관성** 높이기(예: 업무/과제/기술/결정/질문 등)
- 날짜/기한 표준화(예: `YYYY-MM-DD` 또는 “미정”)

### B. Graph로 “자동 탭 생성 + 자동 저장”까지(조건부)
- Azure 평가판 구독 활성화
- Entra 테넌트 생성 및 디렉터리 전환
- 앱 등록 + 공용 클라이언트(Device code) 허용
- 권한: `User.Read`, `Notes.ReadWrite` 동의
- 이후 `onenote_publish.py`를 “섹션(탭) 생성 → 페이지 생성” 흐름으로 확장

