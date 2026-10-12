# IAM Automation Lab

Guided, AI-assisted PowerShell and Microsoft Graph project for a dedicated Microsoft Entra lab tenant. The learner personally ran and verified the live exercises documented below, including selected joiner, mover, and leaver membership actions.

## Verified live results

- Initial read-only inventory: seven users, five enabled, two disabled, zero unknown enabled states.
- Joiner: created and verified the designated fictional lab user, with no assigned license or role from the script.
- Group assignment: created the dedicated assigned security group, added the lab user, verified membership, and reran successfully without duplicate addition.
- Disable: set the designated lab user's accountEnabled flag to false and verified it by reading the user back. Membership remains; existing sessions were not revoked. No actual sign-in attempt was tested.
- Directory role report: three direct active assignments exported. Group members and PIM eligibility are outside coverage.
- MFA registration report: five users returned, all registered. This does not establish enforcement; disabled users are outside this API coverage and report data may lag.
- Stale-account review: eight users exported with a 90-day threshold, six recent successful sign-ins, two unknown history, zero stale candidates. Missing history is not classified as inactivity.
- Final inventory: eight users, five enabled, three disabled, zero unknown enabled states.
- Leaver membership cleanup: removed the disabled lab user from the dedicated lab group, verified absence, and verified that the user remains present and disabled. This is a selected leaver action, not full offboarding or session revocation.
- Repeat leaver cleanup: detected membership already absent, performed no removal, and verified absence and the retained disabled user again.
- Mover: created dedicated Team A and Team B groups, staged the disabled user in A, transferred membership by removing A before adding B, and verified A absent/B present. Repeat Move detected the desired state without duplicate changes. Cleanup verified both memberships absent. The user remained disabled throughout; this demonstrates membership administration, not application access.

Screenshots are indexed in [evidence/README.md](evidence/README.md). Raw reports and local object references are excluded from publication.

## Exercise scope

1. Confirm PowerShell prerequisites and tenant identity.
2. Export a read-only lab user inventory.
3. Create a fictional lab user and assign a designated lab group.
4. Disable that lab user and verify the enabled flag by read-back.
5. Report privileged role assignments, including scope and assignment type where available.
6. Report MFA registration separately from policy enforcement.
7. Export possible stale accounts with explicit handling for missing or unavailable sign-in data.
8. Automate selected joiner, mover, and leaver membership actions with preview, verification, repeat handling, and mover cleanup.

## Validation and reproduction

The original eight operational scripts and the new mover script passed syntax checks. Thirteen offline synthetic safeguard tests and eleven offline mover tests passed; run `Test-LabSafeguards.ps1` and `Test-LabMover.ps1` to reproduce them. These tests do not authenticate or change tenant objects. Live results are documented separately in the evidence index.

See [reproduce.md](reproduce.md) for the exercise sequence, permissions, prerequisites, and limitations. This project covers selected JML actions, not complete offboarding or an HR-driven lifecycle engine.

## Working boundaries

Use the dedicated lab tenant and fictional test users. Check the connected tenant before any operation. Begin with read permissions; request write permissions only when required for a specific exercise. Changes must target explicitly designated lab objects. Reports are local and may contain tenant identifiers or user details; review and sanitize before publishing. Never publish credentials, tokens, authentication codes, or raw sign-in logs.

Some registration and sign-in reports depend on licenses and administrative roles. Missing data or an unavailable API must be reported as unknown or unavailable, not as proof of missing MFA or inactivity. Exact requirements will be verified against Microsoft documentation as each exercise is implemented.

## Sources

- [Microsoft Graph PowerShell getting started](https://learn.microsoft.com/powershell/microsoftgraph/get-started)
- [Microsoft Graph authentication commands](https://learn.microsoft.com/powershell/microsoftgraph/authentication-commands)
