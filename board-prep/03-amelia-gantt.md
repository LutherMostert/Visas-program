# MV Amelia litigation timeline — one-pager

Fill exact dates from Danie Malherbe before circulating. Durations below are placeholder working assumptions — do not present as firm.

## Mermaid Gantt (paste into any mermaid-aware renderer, or screenshot for the board pack)

```mermaid
gantt
    title MV Amelia / Mabasen — litigation timeline
    dateFormat  YYYY-MM-DD
    axisFormat  %b %Y

    section Pre-filing
    Cession perfection / registration   :active, cess, 2026-04-21, 14d
    Demand letter & without-prejudice    :demand, after cess, 10d
    Settlement window (pre-filing)       :crit, settle1, after demand, 14d

    section Filing & pleadings
    Summons issued                       :milestone, file, after settle1, 0d
    Service on Mabasen                   :serve, after file, 7d
    Answer / defence filed               :ans, after serve, 30d
    Replication                          :rep, after ans, 21d

    section Interlocutory
    Discovery                            :disc, after rep, 45d
    Pre-trial conference                 :ptc, after disc, 14d

    section Arrest track (parallel)
    Arrest-readiness pack (WKH)          :crit, arrestprep, 2026-04-21, 7d
    Arrest-trigger window (vessel moves) :crit, arrestwin, 2026-04-28, 180d
    Arrest application (if triggered)    :milestone, arrest, after arrestwin, 0d

    section Resolution
    Settlement window (post-discovery)   :crit, settle2, after ptc, 30d
    Trial                                :trial, after settle2, 21d
    Judgment                             :milestone, judg, after trial, 0d
    Execution / vessel sale              :exec, after judg, 60d
```

## Key milestones for the board to track

| Milestone | Owner | Board visibility |
|---|---|---|
| Cession perfected | Danie | Report at every board until closed |
| Demand letter served | WKH | Update next board |
| Summons filed | WKH | Update next board |
| Answer received | WKH | Dictates next strategy inflection |
| Settlement window 1 (pre-filing) | CEO + WKH | Board sign-off on settlement floor |
| Arrest trigger | CEO + Danie | **Pre-authorised** per board decision |
| Discovery complete | WKH | Provisioning review |
| Settlement window 2 (post-discovery) | CEO + WKH | Board ratify |
| Trial / judgment | WKH | Final resolution reporting |

## Critical path

Cession → Arrest-readiness → Settlement windows.

The two settlement windows are where most cases resolve; the trial track is insurance, not the plan. The arrest track runs in parallel and is the primary source of negotiating leverage — collapses the moment the vessel leaves jurisdiction.

## Dependencies outside our control

- Mabasen's counsel response times (can add 30–60 days at answer / discovery)
- Court roll availability for trial date (NAM High Court — typically 6–9 months out from PTC)
- Vessel movement — outside our operational control; dictates whether arrest track stays live

## What this changes for today

If cession is **not** perfected, weeks 1–2 of the timeline collapse into weeks 1–2 of today. That's the number one ask from Danie this morning.
