import datetime


def coerce_sheet_date_value(value):
    """Normalizes Google Sheets date values into a ``datetime.date``."""
    if value in (None, ''):
        return None

    if isinstance(value, datetime.datetime):
        return value.date()

    if isinstance(value, datetime.date):
        return value

    if isinstance(value, (int, float)):
        serial_start = datetime.date(1899, 12, 30)
        return serial_start + datetime.timedelta(days=float(value))

    if isinstance(value, str):
        normalized = value.strip()
        if not normalized:
            return None
        for fmt in ('%Y-%m-%d', '%m/%d/%Y', '%m/%d/%y'):
            try:
                return datetime.datetime.strptime(normalized, fmt).date()
            except ValueError:
                continue
        try:
            return datetime.date.fromisoformat(normalized)
        except ValueError:
            return None

    return None


def extract_existing_ids(values):
    """Builds a normalized transaction-id set from a Sheets value range."""
    return {
        str(row[0]).strip()
        for row in values
        if row and str(row[0]).strip()
    }


def determine_sync_window(latest_date_str, today=None, lookback_days=730, overlap_days=1):
    """Returns a deterministic sync window from the latest known sheet date."""
    today = today or datetime.date.today()
    normalized_latest = coerce_sheet_date_value(latest_date_str)
    if normalized_latest:
        return (
            normalized_latest - datetime.timedelta(days=overlap_days),
            today,
            'incremental'
        )

    return (
        today - datetime.timedelta(days=lookback_days),
        today,
        'bootstrap'
    )
