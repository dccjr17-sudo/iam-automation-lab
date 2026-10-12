# Reproduce the lab

Use PowerShell 7 and Microsoft.Graph.Authentication in a dedicated Microsoft Entra lab tenant. Authenticate as its administrator using interactive sign-in. Scripts use delegated permissions scoped to their process and disconnect on exit. Device-code authentication is available with `-UseDeviceCode`; never share codes or credentials.

Supply your tenant GUID with `-ExpectedTenantId`. User creation also requires your verified `onmicrosoft.com` domain with `-LabDomain`. Do not use the author's tenant identifiers in a different tenant.

## Exercise sequence

1. Run `Export-LabUsers.ps1` for the baseline inventory.
2. Run `New-LabUser.ps1` without `-Apply` to preview; then with `-Apply` to create the designated lab user. The script refuses to adopt an existing sign-in name. The random initial password is not displayed or persisted; this exercise does not sign in as the new user.
3. Preview and apply `Set-LabGroupMembership.ps1`. It creates the dedicated group and verifies membership. Repeat apply to confirm existing membership is detected.
4. Preview and apply `Disable-LabUser.ps1`. Read-back verifies the disabled flag; this is not a sign-in test or session revocation.
5. Run `Export-LabRoleAssignments.ps1`, `Export-LabMfaRegistration.ps1`, and `Export-LabStaleAccounts.ps1` for local reports. The stale threshold defaults to 90 days and can be changed with `-InactiveDays`.
6. Run the inventory again to compare counts.
7. Preview and apply `Remove-LabGroupMembership.ps1`. The user must already be disabled. The deletion targets the membership reference, retaining the user and group. Repeat to confirm an absent membership causes no deletion.

Creation and membership scripts write tenant-bound references into `runtime/`. Keep those local. Do not edit them to target unrelated objects. If creation succeeded but verification or reference persistence failed, inspect the exact lab object before retrying; do not bypass safeguards by automatically adopting it.

## Permissions requested

| Operation | Delegated Graph scopes |
| --- | --- |
| Inventory | User.Read.All |
| User creation | User.Create, User.Read.All |
| Group creation and membership | Group.ReadWrite.All, User.Read.All |
| Disable user | User.EnableDisableAccount.All, User.Read.All |
| Active directory role report | RoleManagement.Read.Directory |
| MFA registration report | AuditLog.Read.All |
| Stale-account review | User.Read.All, AuditLog.Read.All |
| Membership cleanup | GroupMember.ReadWrite.All, Group.Read.All, User.Read.All |

Membership cleanup uses Group.Read.All to validate group properties and markers as well as the membership write scope. Administrator consent and appropriate caller roles are required. Report availability also depends on licensing; signInActivity requires Entra P1 or P2.

## Offline tests

Run `./Test-LabSafeguards.ps1` from PowerShell 7. Its 13 tests use synthetic responses and a temporary directory. They never connect to Graph, read your live runtime references, or alter tenant objects. The tests cover selected safeguards; live screenshots provide separate evidence of product behavior.

## Limits

This project automates selected joiner, mover, and leaver actions. The mover preparation, transfer, repeat handling, and cleanup were validated live and documented in the evidence index. It is not a complete HR-driven JML engine, production offboarding system, or continuous monitoring service. No scheduler, rollback transaction, immutable audit service, session revocation, or complete inherited-access removal is included.

## Mover membership exercise

Use `Move-LabUser.ps1 -ExpectedTenantId <tenant-guid> -Action Prepare` to preview two dedicated assigned security groups: LAB-IAM-Automation-Team-A and LAB-IAM-Automation-Team-B. Add `-Apply` to create them and stage the designated user in Team A. The user must remain disabled; this demonstrates membership administration, not working application access or a production employee move.

Next preview and apply `-Action Move`. It removes and verifies Team A membership before adding and verifying Team B. Repeat Move to check reconciliation without duplicate writes. Then preview and apply `-Action Cleanup` to remove both memberships and retain the disabled user and groups. Preparation is not a reset; after cleanup, Move refuses to restage access.

Prepare requires User.Read.All and Group.ReadWrite.All. Move and Cleanup require User.Read.All, Group.Read.All, and GroupMember.ReadWrite.All. Saved group references and preparation state stay under runtime and are not published. Existing groups without references are never automatically adopted.

Each membership change is a separate Graph request. A failure after source removal can leave neither membership present; an ambiguous addition can have succeeded remotely. Inspect both groups, then rerun Move to reconcile before cleanup. There is no transaction or automatic restoration of the old membership. Do not run concurrent lifecycle operations on the same objects.

Run `./Test-LabMover.ps1` for eleven synthetic tests covering preparation, validation, move order, repeat runs, interrupted moves, removal verification, and cleanup.

Role reports cover direct active assignments, not expanded group members or PIM eligibility. MFA registration is not enforcement. Disabled users are outside registration report coverage and report data can lag. Missing successful sign-in history is unknown; stale candidates require review. Raw CSV reports contain identity details and stay local.

## Official references

- [Create user](https://learn.microsoft.com/en-us/graph/api/user-post-users?view=graph-rest-1.0)
- [Update user](https://learn.microsoft.com/en-us/graph/api/user-update?view=graph-rest-1.0)
- [Create group](https://learn.microsoft.com/en-us/graph/api/group-post-groups?view=graph-rest-1.0)
- [Add member](https://learn.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-1.0)
- [Remove member reference](https://learn.microsoft.com/en-us/graph/api/group-delete-members?view=graph-rest-1.0)
- [List role assignments](https://learn.microsoft.com/en-us/graph/api/rbacapplication-list-roleassignments?view=graph-rest-1.0)
- [MFA registration report](https://learn.microsoft.com/en-us/graph/api/authenticationmethodsroot-list-userregistrationdetails?view=graph-rest-1.0)
- [List users and sign-in activity](https://learn.microsoft.com/en-us/graph/api/user-list?view=graph-rest-1.0)
