# ExchangeOnlineInventory

Self-contained PowerShell scripts that inventory Exchange Online mailboxes and
export the results to CSV. Two scripts are provided:

- **`ExchangeOnlineInventory.ps1`** — inventories all **active** mailboxes (user,
  shared, room, equipment, etc.) and their in-place archives.
- **`InactiveMailboxInventory.ps1`** — inventories **inactive** mailboxes
  (mailboxes of deleted users preserved by a hold), including the reason each is
  held and the date it became inactive. See
  [Inactive mailbox inventory](#inactive-mailbox-inventory).

For every active mailbox `ExchangeOnlineInventory.ps1` collects:

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

## Inactive mailbox inventory

`InactiveMailboxInventory.ps1` is a **separate** script for **inactive mailboxes** —
the mailboxes of deleted users that are preserved by a Litigation Hold, an
eDiscovery or In-Place Hold, a Microsoft Purview retention policy/label, or a delay
hold. These mailboxes are **not** returned by `ExchangeOnlineInventory.ps1`, so run
this script if you need to account for them (for example, to compare against a
3rd party archive of a former employee's mailbox).

It collects the same size/count/date detail as the main script, plus hold and
identity information:

| Column                  | Description                                                                                                    |
| ----------------------- | -------------------------------------------------------------------------------------------------------------- |
| `Mailbox`               | Primary SMTP address the mailbox had **before** deletion                                                       |
| `DisplayName`           | Display name of the former user                                                                                |
| `MailboxType`           | `Primary` or `Archive`                                                                                         |
| `ExchangeGuid`          | Unique mailbox identifier (used for all lookups)                                                               |
| `DistinguishedName`     | Unique directory identifier                                                                                    |
| `HoldReasons`           | Why the mailbox is held: Litigation / eDiscovery / In-Place / Retention Policy / Retention Label / Delay Hold  |
| `InPlaceHolds`          | Raw In-Place / retention hold identifiers (for audit)                                                          |
| `LitigationHoldEnabled` | `True` / `False`                                                                                              |
| `BecameInactive`        | Date the mailbox became inactive (`WhenSoftDeleted`)                                                           |
| `Messages`              | Total item count                                                                                               |
| `Size (bytes)`          | Total mailbox size in bytes                                                                                    |
| `Size (GB)`             | Total mailbox size in gigabytes (rounded to 2 decimals)                                                        |
| `Oldest Message`        | Received date of the oldest item across all folders                                                            |
| `Newest Message`        | Received date of the newest item across all folders                                                            |

> **SMTP address collisions:** an inactive mailbox can share its old SMTP address
> with a **new active mailbox** that has since reused that address. The `Mailbox`
> column shows the old address for comparison, but it is **not unique** — use
> `ExchangeGuid` (or `DistinguishedName`) to identify the mailbox. All statistics
> are looked up by `ExchangeGuid`, so the figures always refer to the correct
> inactive mailbox regardless of any collision.

> **Org-wide retention policies** do not stamp the `InPlaceHolds` property. If a
> mailbox is inactive solely because of an organisation-wide retention policy,
> `HoldReasons` reads "None detected…"; check `Get-OrganizationConfig | FL InPlaceHolds`.

Prerequisites, installation, throttling behaviour, the `-OutputPath` /
`-MaxRetries` / `-ThrottleDelayMs` parameters, and the `*_failures.log` output all
work exactly as for the main script. Run it the same way:

```powershell
.\InactiveMailboxInventory.ps1
```

The default output file is
`InactiveMailboxInventory_<OrganisationName>_<yyyyMMdd_HHmmss>.csv`. A
**Compliance Administrator** or **Global Reader** role is recommended so all hold
properties are readable.

---

## Author

**Matthew Levy (MVP)**

## License

See [LICENSE](LICENSE).
