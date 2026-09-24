# ExchangeOnlineInventory

A single, self-contained PowerShell script that inventories **all** Exchange Online
mailboxes (user, shared, room, equipment, etc.) and their in-place archives, and
exports the results to a CSV file.

For every mailbox the script collects:

| Column           | Description                                             |
| ---------------- | ------------------------------------------------------- |
| `Mailbox`        | Primary SMTP address of the mailbox                     |
| `MailboxType`    | `Primary` or `Archive`                                  |
| `Messages`       | Total item count (`ItemCount`)                          |
| `Size (bytes)`   | Total mailbox size in bytes                             |
| `Size (GB)`      | Total mailbox size in gigabytes (rounded to 2 decimals) |
| `Oldest Message` | Received date of the oldest item across all folders     |
| `Newest Message` | Received date of the newest item across all folders     |

Each mailbox produces one `Primary` row. If the mailbox has an active in-place
archive, a second `Archive` row is written with the same `Mailbox` value so the
two can be matched. The output is suitable for capacity reporting, audits, or
reconciliation against a third-party mail-archiving report.

> The script is **read-only**. It does not modify, move, or delete any mailbox or
> message.

---

## ⚠️ Important: large tenants and Exchange Online throttling

**Read this before you run the script on a production tenant.**

The script reads folder-level statistics for **every mailbox and every archive**,
which is a heavy reporting workload. Exchange Online applies **dynamic throttling**
to protect the service, so on tenants with many mailboxes you should expect:

- **Long run times** — potentially hours on tenants with tens of thousands of
  mailboxes. This is normal; a progress bar shows how far along it is.
- **Throttling / timeouts** — the service may inject delays or return transient
  errors under sustained load.

The script is built to handle this:

- It **automatically retries** transient throttling, timeout, and connection
  errors with an increasing (exponential) back-off delay, honouring any
  server-suggested wait time.
- It **re-establishes the Exchange Online session** if it drops mid-run.
- Any mailbox that still cannot be read after all retries is recorded in a
  companion **`*_failures.log`** file next to the CSV, and the run continues.
  **Always check that log after a run so you know the report is complete.**

For very large tenants, pace the script with **`-ThrottleDelayMs`** (a short pause
after each mailbox) to reduce the chance of being throttled in the first place, and
tune **`-MaxRetries`** if needed. See [Parameters](#parameters).

> Note: because the script signs in interactively, a multi-hour run could, in rare
> cases, prompt to re-authenticate if the access token fully expires. Keep an eye
> on very long unattended runs.

---

## Prerequisites

1. **Windows** with **PowerShell 5.1** or **PowerShell 7+**.

2. A Microsoft 365 / Exchange Online account with at least the
   **View-Only Recipients** role. The **Global Reader**, **Exchange
   Administrator**, or **Global Administrator** roles also work. The account is
   only used to *read* mailbox statistics.

3. The **Exchange Online Management** PowerShell module.

---

## Installation

Open PowerShell and install the Exchange Online Management module (once per
machine):

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

If prompted to trust the PSGallery repository, answer **Yes** (`Y`).

If the module is already installed, make sure it is current:

```powershell
Update-Module ExchangeOnlineManagement
```

Then download `ExchangeOnlineInventory.ps1` from this repository and save it
somewhere convenient, e.g. `C:\Temp\ExchangeOnlineInventory.ps1`.

---

## Running the script

1. Open PowerShell and change to the folder containing the script:

   ```powershell
   cd C:\Temp
   ```

2. If script execution is blocked, allow it **for this session only**:

   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```

3. Run the script. A browser sign-in window opens (multi-factor authentication is
   supported):

   ```powershell
   .\ExchangeOnlineInventory.ps1
   ```

4. When it finishes, the script prints the full path to the generated CSV file.

By default the CSV is written to the current directory as:

```
ExchangeOnlineInventory_<OrganisationName>_<yyyyMMdd_HHmmss>.csv
```

### Examples

Write to a specific path:

```powershell
.\ExchangeOnlineInventory.ps1 -OutputPath "C:\Reports\EXO_Inventory.csv"
```

Pace the script on a large tenant (200 ms pause after each mailbox):

```powershell
.\ExchangeOnlineInventory.ps1 -ThrottleDelayMs 200
```

---

## Parameters

| Parameter          | Default          | Description                                                                                                                                                     |
| ------------------ | ---------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `-OutputPath`      | timestamped file | Full path (including file name) for the CSV output. If omitted, a timestamped file named after the organisation is created in the current directory.           |
| `-MaxRetries`      | `5`              | Maximum automatic retries per failed statistics call when a transient (throttling/timeout/connection) error occurs. Set to `0` to disable retries. Range 0–20. |
| `-ThrottleDelayMs` | `0`              | Optional pause, in milliseconds, after each mailbox is processed. Use on large tenants to avoid triggering dynamic throttling. Range 0–60000.                  |

---

## Output

- **`<name>.csv`** — the inventory, one row per mailbox (plus one per active
  archive).
- **`<name>_failures.log`** — written **only if** one or more mailboxes could not
  be read after all retries. Lists the timestamp, mailbox, type, and error for
  each failure so you can confirm the report's completeness.

---

## Author

**Matthew Levy (MVP)**

## License

See [LICENSE](LICENSE).
