# GAS Smart Scheduler

A small Google Apps Script library + example to run hybrid schedules:
- runs at short interval during peak hours (e.g. every 10 minutes from 06:00–22:59)
- runs at longer interval during off-peak hours (e.g. once per hour between 23:00–05:59)

Includes:
- `TriggerScheduler` library for creating hybrid triggers
- `pushDataToBackend()` example (sheet → backend) hardened with sheet locking, retry logic, logging, and `_sync_timestamp` injection

Usage:
1. Import these files into a Google Apps Script project (or use clasp).
2. Edit `CONFIG.BACKEND_URL` and `SHEET_NAME` in `src/pushSync.gs`.
3. Run `setupTriggers()` once to create the hybrid schedule.
4. Use `removeTriggers()` to remove them.
