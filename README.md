**VM Lifecycle Manager** is a robust GUI tool designed to automate the decommissioning of physical and virtual servers. Centralize the cleanup of AD, DNS, SCCM, Citrix, and SCVMM into a single, professional dashboard.

## ⚡ Key Features

* **🖥️ Modern GUI:** Intuitive graphical interface for real-time monitoring and control.
* **🔍 Auto-Discovery:** Automatically identifies machine status across Active Directory and VMM.
* **📦 Batch Operations:**
    * Stop System: Graceful OS shutdown.
    * Logical Clean: Automated cleanup of AD, DNS, SCCM, and Citrix records.
    * Delete VM: Permanent removal from the Hypervisor via SCVMM.
* **🛡️ Enterprise Security:** RBAC integration (AD Group) and secure credential handling.
* **📈 Smart Logging:** Full audit trail of actions with automated SMTP email reporting.
* **🚀 Parallel Processing:** Leverages PowerShell Jobs to run simultaneous tasks.

## 🛠️ Infrastructure Requirements

| Component | Requirement |
| :--- | :--- |
| **OS** | Windows Client/Server (PowerShell 5.1+) |
| **Modules** | ActiveDirectory, VirtualMachineManager, Citrix SDK |
| **Connectivity** | Access to DCs, SCVMM, SCCM Provider & Citrix Controllers |
| **Permissions** | Least Privilege access for Decommissioning tasks |

## 🚀 Setup & Execution

### 1. Configuration
Create a DecomConfig.xml file in the project root. Use the provided template, but never commit real production data to the repository.

### 2. Execution
Open the PowerShell console as Administrator and run:
.\VM_LIFECYCLE_MANAGER.ps1

### 3. Workflow
* Login: Authenticate with an account belonging to the allowed AD group.
* Discover: Enter hostnames (comma-separated) and click AUTO-DISCOVER.
* Action: Select the desired operation (STOP, CLEAN, or DELETE) based on the returned status.

## 📂 Project Structure
* VM_LIFECYCLE_MANAGER.ps1: Main script (GUI Interface & Control Logic).
* DecomWorker.ps1: Processing engine that executes background jobs.
* DecomConfig.xml: Configuration file (Template included).
