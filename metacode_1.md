# LGH MODEL — METACODE
## Version: 0.2 — Draft
## Date: May 2026

---

## SECTION 1: AGENTS

| Code | Name | Role |
|------|------|------|
| HoDME | Ryan | Head of Development, Modelling & Execution |
| HoBIF | Charity | Head of Business Infrastructure & Finance |
| HoHM | Chris | Head of Hospitality Management |
| CFO | Zandra | Chief Financial Officer |
| FA | Terry | Financial Accountant |
| MA | Daniel | Management Accountant |
| IM | Thomas | Investment Manager |
| FO | Danielle | Family Office |
| FM | Kevin | Financial Modeller |
| TM | TBD | Technology Manager |
| Prog | TBD | Programmer |
| BA | Pilar | Business Administrator |
| LC | TBD | Legal Counsel |
MAM, sk, metacode administrator

---

## SECTION 2: DATA

| Code | Description | Format | Periodicity | Owner |
|------|-------------|--------|-------------|-------|
| lghQR | LGH quarterly report | PDF | Quarterly | CFO |
| [AF] | Set of audited financials | PDF | Annual | CFO |
| lghAF | LGH audited financials (subset of [AF]) | PDF | Annual | CFO |
| [propAF] | Set of property audited financials (subset of [AF]) | PDF | Annual | CFO |
| [CS2i] | CS2 inputs | Excel | As needed | IM |
| [CS2o] | CS2 outputs | Excel | As needed | IM |
| [QRi] | Quarterly report inputs (subset of [CS2o]) | Excel | Quarterly | IM |
| [AFi] | Audited financials inputs (subset of [CS2o]) | Excel | Annual | IM |
| [PF] | Property forecasts (subset of [CS2i]) | Excel | As needed | MA |
| [DU] | Debt updates (subset of [CS2i]) | Excel | As needed | MA |
| [CPI] | Construction project updates (subset of [CS2i]) | Excel | As needed | FM |
| [S&Ui] | Sources and uses inputs | Excel | As needed | IM |
| [S&Uo] | Sources and uses outputs | Excel | As needed | IM |
| [SKest] | SK monthly estimated cash requirements | Excel | Monthly | FO |
| SKact | SK actual requirement for current month | Excel | Monthly | FO |

---

## SECTION 3: SYSTEMS

| Code | Name | Type | Description |
|------|------|------|-------------|
| S&U | Sources and Uses | Excel | Monthly liquidity model |
| [models] | Set of core financial models | Excel/Python | LGH, C12, CS2 models |
| LGH | LGH model (in [models]) | Excel/Python | LGH investment vehicle model |
| C12 | C12 model (in [models]) | Excel/Python | C12 operating company model |
| [CS2] | Set of CS2 property models (in [models]) | Excel/Python | One per property |
| Opera | Oracle hotel management system | SQL | Property operating data |
| Salesforce | Salesforce CRM | Cloud | Customer and sales data |

---

## SECTION 4: MILESTONES

| Code | Description | Timing |
|------|-------------|--------|
| PropAFpublished | All property audited financials published | Annual — date TBD |
| LGHAFpublished | LGH audited financials published | Annual — date TBD |
| lghQRpublished | LGH quarterly report published | Quarterly — date TBD |

---

## SECTION 5: METACODE 
## runs and maintains all financial reporting systems, produces all reports and updates software and metacode as it goes

begin

  run in parallel
  ## basically 5 separate routines which wait for io from each other as needed
    modelling group routine
    finance group routine
    tech group routine
    HM group routine
    metacode maintenance routine

  ## error handling
  if metacode breaks because i/o between routines fails 
  then 
  head of relevant groups responsible for i/o failure report to MAM immediately

  ## expected to be good for 5yrs, with a major systems upgrade in 2030 prior to launching marketing of REIT for which additional public fund requirements may dictate systems
  until end of 2030 or later if decided good enough, otherwise stop and wait for upgrade

end

### 5.1 tech group routine




### 5.2 tech group routine
## MAM overall responsible but only for initial systems, consulting as needed, and iferror occurs
## HoBIF overall responsibility
## TM responsible day to day
## prog writes all Python codes
## other resources as needed at approvaal by MAM

begin
##initialise systems
draft metacode by end may 2026 and deliver to IM, CFO, HoDME, HoBIF, HoHM, FM
launch production [models] in excel by end may 2026
launch development LGHmodel in python by end may 2026
##update
begin systems update may 2026:

update metacode quarterly
refesh [models] quarterly



subprocedure: refresh
every quarter on first monday of last month of quarter launch refreshed [models]
---

## SECTION 6: SUB-PROCESSES

### SP-01:

```
10
20
30
```

---

*Note: DATA labels to be replaced with actual IO table references once data structure is finalised.*
