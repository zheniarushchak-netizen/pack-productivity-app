import pandas as pd
from io import BytesIO


def select_new_operations(pack, loaded_until):
    operations = pack.copy()

    operations["Operation time"] = pd.to_datetime(
        operations["Operation time"],
        errors="coerce"
    )

    if operations["Operation time"].isna().any():
        raise ValueError(
            "Some operations have missing or invalid timestamps."
        )

    if loaded_until is None:
        new_operations = operations
    else:
        cutoff = pd.Timestamp(loaded_until)

        new_operations = operations.loc[
            operations["Operation time"] > cutoff
        ].copy()
    
    return new_operations

def merge_operations(
    saved_operations,
    uploaded_operations,
    loaded_until
):
    new_operations = select_new_operations(
        uploaded_operations,
        loaded_until
    )

    if saved_operations is None:
        combined_operations = new_operations.copy()
    else:
        combined_operations = pd.concat(
            [saved_operations, new_operations],
            ignore_index=True
        )

    if combined_operations.empty:
        raise ValueError("No operations are available.")

    updated_loaded_until = combined_operations["Operation time"].max()

    return combined_operations, new_operations, updated_loaded_until

def serialize_operations(operations):
    buffer = BytesIO()

    operations.to_parquet(
        buffer,
        engine="pyarrow",
        compression="gzip",
        index=False
    )

    return buffer.getvalue()

def deserialize_operations(operations_bytes):
    operations = pd.read_parquet(
        BytesIO(operations_bytes),
        engine="pyarrow"
    )

    return operations

def validate_upload_period(upload_start, upload_end, loaded_until):
    start = pd.Timestamp(upload_start)
    end = pd.Timestamp(upload_end)

    if pd.isna(start) or pd.isna(end):
        raise ValueError("Upload period contains an invalid timestamp.")

    if start >= end:
        raise ValueError("Upload start must be earlier than upload end.")

    if loaded_until is not None:
        cutoff = pd.Timestamp(loaded_until)

        if pd.isna(cutoff):
            raise ValueError("Saved coverage contains an invalid timestamp.")

        if start > cutoff:
            raise ValueError(
                "There is a gap between saved data and the uploaded period."
            )

    return start, end

def select_upload_period(operations, upload_start, upload_end):
    period_operations = operations.copy()

    times = pd.to_datetime(
        period_operations["Operation time"],
        errors="coerce"
    )

    if times.isna().any():
        raise ValueError(
            "Some operations have missing or invalid timestamps."
        )

    period_operations["Operation time"] = times

    return period_operations.loc[
        (times >= upload_start) & (times < upload_end)
    ].copy()