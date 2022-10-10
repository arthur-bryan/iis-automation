from colorama import Fore, Style
from datetime import datetime
from time import sleep
import subprocess
import sys
import os

TARGETS = ("AW1P-RNNBE-135", "AW1P-RNNBE-136", "AW1P-RNNBE-137")  # hostname of target servers
RECYCLE_DELAY = 5  # delay between recycles (in seconds)
all_pools = []
COMMANDS = {
    "GET_POOLS_STATUS_COMMAND": "Get-IISAppPool  | Format-List Name, State, AutoStart, StartMode",
    "IIS_RESET_COMMAND": "iisreset",
    "START_APP_POOL_COMMAND": "Start-WebAppPool -Name ",
    "STOP_APP_POOL_COMMAND": "Stop-WebAppPool -Name "
}


class IISPool:

    def __init__(self, hostname, name, state, auto_start, start_mode):
        self.hostname = hostname
        self.name = name
        self.state = state
        self.auto_start = auto_start
        self.start_mode = start_mode


def run_command(hostname, command, pool_name=None):
    global COMMANDS, domain_name, user_name, user_password
    if pool_name is not None:
        command = COMMANDS[command] + pool_name
    else:
        command = COMMANDS[command]
    raw_command = f"""
            $host_domain='{domain_name}'
            $user='{user_name}@{domain_name}'
            $pass='{user_password}'
            $secret=ConvertTo-SecureString -AsPlainText ${user_password} -Force
            $cred=New-Object System.Management.Automation.PSCredential -ArgumentList {user_name}@{domain_name},$secret
            Invoke-Command -ScriptBlock {{ {command} }} -ComputerName {hostname}.{domain_name} -Credential $cred 
    """
    print(raw_command)
    output = subprocess.run(["&", "Invoke-Command", "-ScriptBlock", f"{{ {command} }}",
                             "-ComputerName", f"{hostname}.{domain_name}",
                             "-Credential", "New-Object", "System.Management.Automation.PSCredential",
                             "-ArgumentList", "ArgumentList,ConvertTo-SecureString",
                             "-AsPlainText", f"${user_password}", "-Force"],
                            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    return output


# """
# Name      : MobilidadeWS
# State     : Started
# AutoStart : True
# StartMode : AlwaysRunning
#
# Name      : PCG.Portal.WS_4
# State     : Stopped
# AutoStart : False
# StartMode : AlwaysRunning
# """


def output_to_pool_objects(hostname, command_output):
    """converts pools attributes command output to object, then add they to a list"""
    stdout = str(command_output.stdout, 'utf-8', errors='ignore')
    stderr = str(command_output.stderr, 'utf-8', errors='ignore')
    print(stdout, stderr)
    objects = []
    splited_output = stdout.lstrip("\n").rstrip("\n").split("\n\n")
    for output in splited_output:
        iis_pool = IISPool(
            hostname=hostname,
            name=output.split("\n")[0].split(":")[1].strip(),
            state=output.split("\n")[1].split(":")[1].strip(),
            auto_start=output.split("\n")[2].split(":")[1].strip(),
            start_mode=output.split("\n")[3].split(":")[1].strip()
        )
        objects.append(iis_pool)
    return objects


def show_iisreset_output(hostname, command_output):
    # print(f"{iis_pool.hostname.ljust(20, ' ')}{iis_pool.name.ljust(20, ' ')}{iis_pool.state.ljust(20, ' ')}"
    #       f"{iis_pool.auto_start.ljust(20, ' ')}{iis_pool.start_mode.ljust(20, ' ')}{command_output}")
    print(f"{hostname.ljust(20, ' ')}{'Successfuly reseted'.ljust(80, ' ')}")


def show_formatted_pool_attributes(iis_pool):
    print(f"{iis_pool.hostname.ljust(20, ' ')}{iis_pool.name.ljust(20, ' ')}{iis_pool.state.ljust(20, ' ')}"
          f"{iis_pool.auto_start.ljust(20, ' ')}{iis_pool.start_mode.ljust(20, ' ')}")


def main():
    menu_choice = input("""
    \r[ IIS Pools  ]
    \r[ 1 ] Recycle all pools (IIS Reset)
    \r[ 2 ] Start all (disabled) pools
    \r[ 3 ] Stop all (started) pools
    \r[ 4 ] Get all pools status
    \r[ 0 ] Exit
    \r>> """)
    if menu_choice == "1":
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        print(f"{Fore.LIGHTWHITE_EX}[ STARTED AT {current_time} ]{Style.RESET_ALL}")
        print(f"{'SERVER'.ljust(20, ' ')}{'OUTPUT'.ljust(80, ' ')}\n{100 * '-'}")
        for hostname in TARGETS:
            output = run_command(hostname=hostname, command="IIS_RESET_COMMAND")
            show_iisreset_output(hostname=hostname, command_output=output)
            sleep(RECYCLE_DELAY)
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        print(f"{Fore.LIGHTWHITE_EX}[ FINISHED AT {current_time} ]{Style.RESET_ALL}")
    elif menu_choice == "2":
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        print(f"{Fore.LIGHTWHITE_EX}[ STARTED AT {current_time} ]{Style.RESET_ALL}")
        print(f"{'SERVER'.ljust(20, ' ')}{'POOL NAME'.ljust(20, ' ')}{'POOL STATE'.ljust(20, ' ')}"
              f"{'POOL AUTO START'.ljust(20, ' ')}{'POOL START MODE'.ljust(20, ' ')}\n{100 * '-'}")
        for hostname in TARGETS:
            output = run_command(hostname=hostname, command="GET_POOLS_STATUS_COMMAND")
            objects = output_to_pool_objects(hostname=hostname, command_output=output)
            for iis_pool in objects:
                if iis_pool.state != "Started":
                    print(f"Starting '{iis_pool.name}' pool at '{iis_pool.hostname}'...")
                    run_command(hostname=hostname, command="START_APP_POOL_COMMAND", pool_name=iis_pool.name)
            output = run_command(hostname=hostname, command="GET_POOLS_STATUS_COMMAND")
            objects = output_to_pool_objects(hostname=hostname, command_output=output)
            for iis_pool in objects:
                show_formatted_pool_attributes(iis_pool=iis_pool)
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        print(f"{Fore.LIGHTWHITE_EX}[ FINISHED AT {current_time} ]{Style.RESET_ALL}")
    elif menu_choice == "3":
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        print(f"{Fore.LIGHTWHITE_EX}[ STARTED AT {current_time} ]{Style.RESET_ALL}")
        print(f"{'SERVER'.ljust(20, ' ')}{'POOL NAME'.ljust(20, ' ')}{'POOL STATE'.ljust(20, ' ')}"
              f"{'POOL AUTO START'.ljust(20, ' ')}{'POOL START MODE'.ljust(20, ' ')}\n{100 * '-'}")
        for hostname in TARGETS:
            output = run_command(hostname=hostname, command="GET_POOLS_STATUS_COMMAND")
            objects = output_to_pool_objects(hostname=hostname, command_output=output)
            for iis_pool in objects:
                if iis_pool.state == "Started":
                    print(f"Stopping '{iis_pool.name}' pool at '{iis_pool.hostname}'...")
                    run_command(hostname=hostname, command="STOP_APP_POOL_COMMAND", pool_name=iis_pool.name)
            output = run_command(hostname=hostname, command="GET_POOLS_STATUS_COMMAND")
            objects = output_to_pool_objects(hostname=hostname, command_output=output)
            for iis_pool in objects:
                show_formatted_pool_attributes(iis_pool=iis_pool)
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        print(f"{Fore.LIGHTWHITE_EX}[ FINISHED AT {current_time} ]{Style.RESET_ALL}")
    elif menu_choice == "4":
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        print(f"{Fore.LIGHTWHITE_EX}[ STARTED AT {current_time} ]{Style.RESET_ALL}")
        print(f"{'SERVER'.ljust(20, ' ')}{'POOL NAME'.ljust(20, ' ')}{'POOL STATE'.ljust(20, ' ')}"
              f"{'POOL AUTO START'.ljust(20, ' ')}{'POOL START MODE'.ljust(20, ' ')}\n{100 * '-'}")
        for hostname in TARGETS:
            output = run_command(hostname=hostname, command="GET_POOLS_STATUS_COMMAND")
            objects = output_to_pool_objects(hostname=hostname, command_output=output)
            for iis_pool in objects:
                show_formatted_pool_attributes(iis_pool=iis_pool)
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        print(f"{Fore.LIGHTWHITE_EX}[ FINISHED AT {current_time} ]{Style.RESET_ALL}")
    elif menu_choice == "0":
        sys.exit(0)
    else:
        main()


domain_name = ""
user_name = ""
user_password = ""
try:
    domain_name = os.environ["VM_DOMAIN"]
    user_name = os.environ["USER_NAME"]
    user_password = os.environ["USER_PASSWORD"]
except KeyError:
    print("Please, set the credential environment variables!")
    sys.exit(1)

if __name__ == "__main__":
    main()
