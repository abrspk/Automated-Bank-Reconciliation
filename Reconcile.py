import pandas as pd


def reconcile_bank_and_ledger(bank_file, ledger_file, output_file):
    bank_df = pd.read_excel(bank_file)
    ledger_df = pd.read_excel(ledger_file)

    bank_df["Reference_ID"] = bank_df["Reference_ID"].astype(str).str.strip()
    ledger_df["Reference_ID"] = ledger_df["Reference_ID"].astype(str).str.strip()

    merged = pd.merge(
        bank_df,
        ledger_df,
        on="Reference_ID",
        how="outer",
        suffixes=("_Bank", "_Ledger"),
        indicator=True,
    )

    matched = merged[
        (merged["_merge"] == "both")
        & (merged["Amount_PKR_Bank"] == merged["Amount_PKR_Ledger"])
    ].copy()

    amount_mismatch = merged[
        (merged["_merge"] == "both")
        & (merged["Amount_PKR_Bank"] != merged["Amount_PKR_Ledger"])
    ].copy()

    only_in_bank = merged[merged["_merge"] == "left_only"].copy()
    only_in_ledger = merged[merged["_merge"] == "right_only"].copy()

    summary = pd.DataFrame(
        {
            "Category": [
                "Matched Transactions",
                "Amount Mismatch",
                "Unrecorded in Ledger (Bank Only)",
                "Unpresented in Bank (Ledger Only)",
            ],
            "Count": [
                len(matched),
                len(amount_mismatch),
                len(only_in_bank),
                len(only_in_ledger),
            ],
        }
    )

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Reconciliation_Summary", index=False)
        matched.drop(columns=["_merge"]).to_excel(writer, sheet_name="Matched", index=False)
        amount_mismatch.drop(columns=["_merge"]).to_excel(
            writer, sheet_name="Amount_Mismatch", index=False
        )
        only_in_bank.drop(columns=["_merge"]).to_excel(
            writer, sheet_name="Bank_Only_Entries", index=False
        )
        only_in_ledger.drop(columns=["_merge"]).to_excel(
            writer, sheet_name="Ledger_Only_Entries", index=False
        )

    print(f"Reconciliation completed successfully. Report saved to {output_file}")


if __name__ == "__main__":
    reconcile_bank_and_ledger(
        "sample_bank_and_ledger.xlsx",
        "sample_bank_and_ledger.xlsx",
        "Reconciliation_Audit_Report.xlsx",
    )
