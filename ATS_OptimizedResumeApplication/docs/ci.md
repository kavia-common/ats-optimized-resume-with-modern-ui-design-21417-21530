Excel must run on a Windows host. The container provides file-based packaging/validation only.
Do NOT install pywin32 inside this Linux container. To use pywin32, run tests on a Windows host and set INSTALL_PYWIN32=1 there.
To embed a real vbaProject.bin: open a valid .xlsm on Windows, extract vbaProject.bin and place it in templates/ replacing the placeholder.
Environment variables:
  - INSTALL_OFFICEJS=1 to install office-addin-cli tooling (optional)
  - OFFICE_ADDIN_CLI_VERSION to pin office-addin-cli version (optional)
  - OFFICE_ADDIN_GLOBAL=1 to perform a global npm -g install (optional; default avoids global install)
  - INSTALL_PYWIN32=1 to install pywin32 (Windows host only)
  - UPGRADE_PIP=1 to upgrade pip (not recommended by default)
  - CONTAINER_WORKSPACE to override default workspace path
