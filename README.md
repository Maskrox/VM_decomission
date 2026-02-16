# VM Lifecycle Manager (VLM) - Enterprise Edition

A robust PowerShell-based GUI tool for automated server decommissioning. 
This utility manages the end-of-life process for virtual and physical servers by coordinating actions across Active Directory, VMM, SCCM, and DNS.

## Key Features

- **Batch Processing:** Support for multiple hostnames via comma-separated input.
- **Environment Discovery:** Automatic identification of VM placement across multiple VMM servers.
- **Safety Barriers:** - Ping-check verification before logical cleanup.
  - "Safety Word" confirmation for destructive hypervisor actions.
  - Physical hardware protection (prevents disk deletion on physical nodes).
- **XML-Driven:** No sensitive data hardcoded in the script; infrastructure is managed via external configuration.
- **Audit Ready:** Integrated logging with Task/Ticket number prefixes for traceability.

## Prerequisites

- **Modules:** `ActiveDirectory`, `VirtualMachineManager`, `ConfigurationManager`.
- **Permissions:** Account must have delegated rights to delete objects in AD, SCCM, and the Hypervisor.
- **Connectivity:** Line of sight to Domain Controllers, VMM Servers, and SCCM Providers.

Configure Infrastructure:
Locate DecomConfig_Example.xml in the root folder.
Rename it to DecomConfig.xml.
Update the XML with your specific VMM hosts, Domain Controllers, and Target OUs.

Run the tool:
Execute DecomManager_v6.1.ps1 in a PowerShell terminal or ISE.
The tool will attempt to auto-load DecomConfig.xml on startup.
