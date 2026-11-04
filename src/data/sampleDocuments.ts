export interface SampleDocument {
  id: string;
  title: string;
  category: string;
  description: string;
  content: string;
}

export const SAMPLE_DOCUMENTS: SampleDocument[] = [
  {
    id: 'sample-1',
    title: '클라우드 네이티브 MSA 전환 2차 아키텍처 회의',
    category: '인프라/백엔드',
    description: '모놀리식에서 마이크로서비스로 전환하며 gRPC, 서비스 메시(Istio), 멱등성 보장 방안을 검토하는 회의록',
    content: `[회의명] 모놀리식 커머스 플랫폼의 MSA(Microservices Architecture) 2차 전환 검토
[일시] 2026-11-04 14:00 ~ 15:30 (대회의실 A)
[참석자] 인프라팀장, 백엔드 파트장, 결제 모듈 담당자, 데브옵스 엔지니어

1. 회의 목적
기존 단일 애플리케이션(Monolithic) 환경에서 주문 폭주 시 발생하던 DB 데드락과 단일 장애점(SPOF)을 해소하기 위해, 주문/결제/회원 도메인을 독립 서비스로 분리하고 클라우드 네이티브 기반 서비스 메시(Service Mesh) 및 비동기 이벤트 파이프라인을 도입하기 위함.

2. 주요 논의 안건
(1) 서비스 간 통신 프로토콜 표준화:
- RESTful HTTP/1.1 API의 오버헤드를 줄이기 위해 내부 마이크로서비스 간 통신은 Protocol Buffers 기반의 gRPC(HTTP/2)를 채택하기로 합의함.
- 외부 모바일/웹 클라이언트와의 접점에는 BFF(Backend For Frontend) 패턴 및 API Gateway(Kong)를 배치하여 라우팅 및 JWT 토큰 인가를 단일화함.

(2) 서비스 메시(Service Mesh) 도입:
- 쿠버네티스(Kubernetes) 클러스터 내 서비스 간 mTLS(상호 TLS) 암호화 통신 및 카나리(Canary) 배포 트래픽 제어를 위해 Istio 사이드카 프록시(Envoy)를 적용할 계획임.
- 인프라팀에서는 Envoy 사이드카 인젝션에 따른 파드 메모리 추가 사용량(약 50MB/pod)에 대한 리소스 프로비저닝 예산을 산정 중임.

(3) 분산 트랜잭션과 데이터 정합성:
- 결제 완료 후 재고 차감 및 배송 지시 단계에서 2PC(Two-Phase Commit)의 락킹 병목을 피하기 위해, 사가(Saga) 패턴(코레오그래피 방식)을 채택함.
- 메시지 브로커로 Apache Kafka를 사용하여 이벤트를 발행(Publish)하며, 네트워크 순단이나 컨슈머 재시도에 대비하여 각 서비스 수신부에 멱등성(Idempotency) 보장 키(UUID)를 저장하고 중복 처리를 차단하는 로직을 필수 구현하기로 함.

(4) 관측 가능성(Observability):
- OpenTelemetry SDK를 백엔드 서비스 전반에 내장하여 분산 추적(Distributed Tracing)용 TraceID를 헤더로 전파하고, Jaeger를 통해 지연 시간(Latency) 병목 구간을 시각화함.

3. 액션 아이템
- 백엔드팀: 다음 주 수요일까지 사가 패턴 실패 시 실행할 보상 트랜잭션(Compensating Transaction) 명세서 초안 작성
- 인프라팀: 스테이징 환경 EKS 클러스터에 Istio 1.24 설치 및 부하 테스트 벤치마크 진행
- 결제팀: PG사 연동 모듈의 결제 승인 웹훅 수신 시 멱등키 캐싱(Redis TTL 24시간) 로직 구현`,
  },
  {
    id: 'sample-2',
    title: '글로벌 결제 PG사 연동 및 정산 시스템 개편 스프린트',
    category: '핀테크/결제',
    description: '신규 PG 연동, 3D Secure 2.0 인증, 에스크로 정산 배치 및 이상금융거래탐지(FDS) 적용 논의',
    content: `[회의명] 해외 결제 PG 연동 및 실시간 정산 마감 시스템 킥오프
[일시] 2026-11-05 10:00 ~ 11:30
[작성자] 핀테크 프로덕트 오너

1. 개요 및 배경
동남아 및 일본 시장 신규 진출을 위해 Stripe 및 현지 로컬 결제 수단(E-Wallet)을 통합 수용하는 글로벌 결제 파이프라인을 구축함. 또한 카드 결제 부정 사용에 대비한 3D Secure 2.0 규격 준수와 일일 정산 대사(Reconciliation) 자동화가 필수적임.

2. 세부 검토 사항
(1) 결제 승인 플로우 및 3DS 2.0:
- 부정 결제 분쟁(Chargeback) 발생 시 가맹점 책임 면제(Liability Shift) 혜택을 받기 위해 Visa/Mastercard의 3-D Secure 2.0 위험 기반 인증(Risk-based Authentication)을 필수 적용함.
- 사용자 이탈률을 낮추기 위해 저위험 트랜잭션은 Frictionless Flow(무인증 통과)로 유도하고, 고위험 시에만 생체/OTP 챌린지를 요청함.

(2) 정산 대사(Settlement Reconciliation) 파이프라인:
- 매일 자정 PG사로부터 수신되는 결제 원장(Raw Clearing Data)과 자사 데이터베이스의 주문 레코드를 1:1 대조하는 배치(Batch) 잡 개발.
- 불일치 건(금액 차이, 승인 취소 누락, 환율 오차)은 별도의 에스크로(Escrow) 보류 계정에 임시 격리 후 재무팀 슬랙 채널로 실시간 알림 발송.

(3) 멱등성(Idempotency) 및 웹훅(Webhook) 처리:
- PG사 통신망 장애로 인한 중복 승인 요청을 방지하기 위해 모든 API 호출에 Client Generated Idempotency-Key를 전송함.
- 웹훅은 비동기 큐(SQS)를 통해 유실 없이 소비하며, Signature 해시 검증(HMAC-SHA256)을 거쳐 변조 여부를 검증함.

3. 미확인 사항 및 주의사항
- 현지 부가가치세(VAT) 원천징수 세율은 현지 법률 대리인의 검토가 아직 완료되지 않음.`,
  },
  {
    id: 'sample-3',
    title: '사내 생성형 AI PoC 결과 보고 및 데이터 보안 가이드 심의',
    category: 'AI/보안',
    description: '사내 문서 검색 RAG 구축 결과와 민감정보 유출 방지(DLP), 프롬프트 인젝션 방어책 심의',
    content: `[회의명] 생성형 AI 기반 사내 지식 관리 시스템 PoC 최종 평가 및 정보보호 심의
[일시] 2026-11-06 15:00
[배석] AI 혁신랩 연구원, 정보보안팀장, 컴플라이언스 변호사

1. 회의 목적
지난 3개월간 수행한 사내 Confluence 및 기술 문서 기반 RAG(Retrieval-Augmented Generation) 시스템의 검색 정확도와 응답 품질을 보고하고, 운영 환경 승격을 위한 보안 통제 가이드라인을 확정함.

2. 주요 보고 내용
(1) RAG 시스템 아키텍처 및 벤치마크 결과:
- 사내 기술 문서 20만 건을 청킹(Chunk Size 512, Overlap 50)하여 벡터 데이터베이스(Milvus)에 Dense Vector로 색인함.
- 검색 단계에서 키워드 기반 BM25(Sparse Vector)와 밀집 벡터 검색을 결합한 하이브리드 검색(Hybrid Search) 및 Re-ranking(Cross-Encoder)을 적용하여 Top-5 문서 적합도(Recall@5)가 78%에서 92%로 향상됨.
- 환각 현상(Hallucination) 억제를 위해 모델 프롬프트에 "컨텍스트에 없는 내용은 반드시 불확실성을 표시할 것"을 명시함.

(2) 정보보호 및 컴플라이언스 요구사항:
- DLP(Data Loss Prevention) 필터를 API 전단에 배치하여 주민등록번호, 계좌번호, 사내 비밀 코드(API Key, DB 연결 문자열)가 프롬프트로 전송될 경우 마스킹(Masking) 처리함.
- 간접 프롬프트 인젝션(Indirect Prompt Injection) 방어를 위해 외부 문서 내 악의적 시스템 지시문 무력화 샌드박스를 구성함.
- 사외 모델 호출 시 제로 데이터 리텐션(Zero Data Retention, ZDR) 계약 체결 여부를 법무팀과 최종 검토함.

3. 의결 사항
- 1단계로 비기밀성 공개 기술 규격서에 한해 사내 베타 오픈 승인.
- 인사/재무 데이터는 권한 기반 접근 제어(RBAC) 연동 완료 후 2차 심의를 거치기로 함.`,
  },
];
