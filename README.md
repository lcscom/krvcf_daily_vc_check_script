# 🛡️ VMware Infrastructure Daily Health Check & Security Audit (v11.0)

> **Multi-vCenter Enterprise Morning Health Check & Administrator Security Audit Engine for PowerShell / PowerCLI**

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://microsoft.com/powershell)
[![PowerCLI](https://img.shields.io/badge/PowerCLI-12.0%2B-vmware.svg)](https://code.vmware.com/web/dp/tool/powercli)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

전 세계 여러 데이터센터에 분산된 VMware vSphere(vCenter, ESXi Host, Datastore, VM) 환경의 **상태를 24/7 추적하고, 매일 아침 Outlook 호환 HTML 보고서를 자동 생성하여 전달하는 엔터프라이즈 자동화 솔루션**입니다.

DNS 쿼리로 인한 성능 병목을 최소화하고, vCenter 내부 디스크 파티션(/storage/log 포함) 꽉 참 예방, 호스트별 가상화 집적도(vCPU:pCore 비율, Memory Oversubscription %) 및 Administrator 계정 보안 감사 기능까지 포함되어 있습니다.

---

## 🌟 Key Features (주요 기능)

### 1. 📢 Global Attention Dashboard (통합 알람)
- 인프라 전체의 크리티컬 경고(물리 NIC Down, Dead Path, vCenter 인증서 만료, 스토리지 용량 부족, 관리자 로그인 실패 등)를 상단에 즉시 요약 표출.

### 2. 📊 VM Inventory Daily Trend Tracking (VM 증감 추적)
- 매일 수집된 VM 목록을 로컬 JSON 데이터베이스(`vm_inventory_history.json`)와 자동 비교.
- 최근 24시간 동안 **신규 배포된 VM(+)** 및 **삭제/미검출된 VM(-)**을 자동 추출하여 변동 현황 제공.

### 3. 🏢 vCenter Appliance Status & Partition Monitoring
- vCenter 어플라이언스 SSL 인증서 만료 D-Day 계산.
- VCSA 내부 주요 디스크 파티션(`/storage/log`, `/storage/db`, `/storage/seat`, `/`)의 용량을 추적하여 로그 폭주로 인한 vCenter 서비스 다운 방지.

### 4. 🖥️ Host Virtualization Density & Physical Link Monitoring
- **CPU 가상화율 (%)**: $\frac{\text{전체 VM vCPU 수}}{\text{호스트 물리 pCore 수}} \times 100$ (400% 초과 시 경고)
- **Memory 가상화율 (%)**: $\frac{\text{전체 VM Memory (GB)}}{\text{호스트 물리 Memory (GB)}} \times 100$ (120% 초과 시 경고)
- 물리 네트워크 포트(`vmnic`)의 **Link Down** 상태 및 Multipath 스토리지 **Dead Path** 추적.

### 5. 🧱 VM Inventory & Integrated Snapshot Management
- VM별 상태, 할당 리소스, 실시간 CPU/Memory 점유율 프로그레스 바(Bar Chart) 표출.
- **Cross-NUMA 구조 감지**: vCPU > 8 및 `CoresPerSocket=1` 구성 VM 자동 탐지.
- **통합 스냅샷 모니터링**: 7일 이상 장기 방치된 스냅샷(Old Snapshots) 및 Disk Consolidation 필요 상태 탐지.

### 6. 🔐 Administrator Security & Activity Audit (보안 감사)
- **시간대별 로그인 트렌드**: 24시간 내 Administrator 로그인 횟수 추적.
- **접속 출처 분석**: 접속 IP 및 DNS 역방향 조회(Cache 적용)를 통한 Client IP/Hostname/App 요약.
- **보안 경고 (Auth Failures)**: 계정 로그인 실패 및 미인가 접근 시도 감지.
- **순수 관리 작업 감사**: VM 생성/삭제/수정, vMotion, 디스크 확장 등 실제 수행된 주요 작업 이력 추적.

---

## 🛠️ Prerequisites (사전 준비 사항)

본 스크립트를 실행하기 위해 아래 환경이 필요합니다.

* **OS**: Windows Server 2016 이상 또는 Windows 10/11 (PowerShell 5.1 이상)
* **PowerCLI**: VMware.PowerCLI 12.0 이상
* **Network**: vCenter Server(Port 443) 및 Gmail/SMTP Server(Port 587) 통신 허용

### PowerCLI 모듈 설치 명령어
```powershell
Install-Module -Name VMware.PowerCLI -Scope CurrentUser -AllowClobber
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
