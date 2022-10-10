from winrm.protocol import Protocol
from classes.credential import DomainCredential
from classes.host import Host


class WindowsManager:

    COMMANDS = {
        "GET_POOLS_STATUS_COMMAND": "Get-IISAppPool  | Format-List Name, State, AutoStart, StartMode",
        "IIS_RESET_COMMAND": "iisreset",
        "START_APP_POOL_COMMAND": "Start-WebAppPool -Name ",
        "STOP_APP_POOL_COMMAND": "Stop-WebAppPool -Name "
    }

    def __init__(self, targets):
        self.__threads_number = 20
        self.targets = targets
        self.target_objects = {}
        self.domains_credentials = {}
        self.transport = "ntlm"
        self.server_cert_validation = "ignore"
        self.__fill_domain_credentials()

    def run_command(self, host, command, pool_name=None):
        p = Protocol(
            endpoint=f"http://{host.ip_address}:5985/wsman",
            transport=self.transport,
            username=rf"{host.credential.domain}\{host.credential.user_name}",
            password=host.credential.user_password,
            server_cert_validation=self.server_cert_validation)
        if pool_name is not None:
            command = self.COMMANDS[command] + pool_name
        else:
            command = self.COMMANDS[command]
        shell_id = p.open_shell()
        command_id = p.run_command(shell_id, 'powershell.exe', ['-command', f'"{command}"'])
        std_out, std_err, status_code = p.get_command_output(shell_id, command_id)
        p.cleanup_command(shell_id, command_id)
        p.close_shell(shell_id)
        return {
            "stdout": std_out.decode('windows-1252'),
            "stderr": std_err.decode('windows-1252'),
            "status_code": status_code
        }

    def __fill_domain_credentials(self):
        for target in self.targets:
            if target["domain_name"] not in self.domains_credentials.keys():
                self.__get_domain_credential(target["domain_name"])
            target_name = target["hostname"] if target["hostname"] else target["ip_address"]
            target_object = Host(target_name, target["ip_address"], self.domains_credentials[target["domain_name"]]["object"])
            self.target_objects.update({target_name: {"object": target_object}})

    def __get_domain_credential(self, domain_name):
        user_name = input(f"[ ? ] User name for {domain_name}: ")
        user_password = input(rf"[ ? ] Password for {domain_name}\{user_name}: ")
        credential = DomainCredential(domain=domain_name, user_name=user_name, user_password=user_password)
        self.domains_credentials.update({credential.domain: {"object": credential}})
