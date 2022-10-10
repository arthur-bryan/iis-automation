from classes.windows_manager import WindowsManager
from classes.credential import DomainCredential
from classes.iis_pool import IISPool
from classes.host import Host
from colorama import Fore, Style
from datetime import datetime
from time import sleep
import targets
import json
import sys
import os

RECYCLE_DELAY = 5   # in seconds


def main():
    manager = WindowsManager(targets.targets)
    while True:
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
            show_output_header()
            for target in manager.target_objects.values():
                result = manager.run_command(target["object"], "IIS_RESET_COMMAND")
                sleep(RECYCLE_DELAY)
                result = manager.run_command(target["object"], "GET_POOLS_STATUS_COMMAND")
                if not result["stdout"] and result["status_code"] != 0:
                    show_get_pools_error(target, result)
                    continue
                objects = output_to_pool_objects(hostname=target["object"].hostname, command_output=result["stdout"])
                for iis_pool in objects:
                    show_formatted_pool_attributes(iis_pool, "IIS_RESET_COMMAND", result["status_code"])
            current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            print(f"{Fore.LIGHTWHITE_EX}[ FINISHED AT {current_time} ]{Style.RESET_ALL}")

        elif menu_choice == "2":
            current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            print(f"{Fore.LIGHTWHITE_EX}[ STARTED AT {current_time} ]{Style.RESET_ALL}")
            show_output_header()
            for target in manager.target_objects.values():
                result = manager.run_command(target["object"], "GET_POOLS_STATUS_COMMAND")
                if not result["stdout"] and result["status_code"] != 0:
                    show_get_pools_error(target, result)
                    continue
                objects = output_to_pool_objects(hostname=target["object"].hostname, command_output=result["stdout"])
                for iis_pool in objects:
                    if iis_pool.state != "Started":
                        # print(f"Starting '{iis_pool.name}' pool at '{iis_pool.hostname}'...")
                        result = manager.run_command(target["object"], "START_APP_POOL_COMMAND", pool_name=iis_pool.name)
                        if not result["stdout"] and result["status_code"] != 0:
                            show_get_pools_error(target, result)
                            continue
                result = manager.run_command(target["object"], "GET_POOLS_STATUS_COMMAND")
                if not result["stdout"]:
                    show_get_pools_error(target, result)
                    continue
                objects = output_to_pool_objects(hostname=target["object"].hostname, command_output=result["stdout"])
                for iis_pool in objects:
                    show_formatted_pool_attributes(iis_pool, "START_APP_POOL_COMMAND", result["status_code"])
            current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            print(f"{Fore.LIGHTWHITE_EX}[ FINISHED AT {current_time} ]{Style.RESET_ALL}")

        elif menu_choice == "3":
            current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            print(f"{Fore.LIGHTWHITE_EX}[ STARTED AT {current_time} ]{Style.RESET_ALL}")
            show_output_header()
            for target in manager.target_objects.values():
                result = manager.run_command(target["object"], "GET_POOLS_STATUS_COMMAND")
                if not result["stdout"] and result["status_code"] != 0:
                    show_get_pools_error(target, result)
                    continue
                objects = output_to_pool_objects(hostname=target["object"].hostname, command_output=result["stdout"])
                for iis_pool in objects:
                    if iis_pool.state == "Started":
                        # print(f"Stopping '{iis_pool.name}' pool at '{iis_pool.hostname}'...")
                        result = manager.run_command(target["object"], command="STOP_APP_POOL_COMMAND", pool_name=iis_pool.name)
                        if not result["stdout"] and result["status_code"] != 0:
                            show_get_pools_error(target, result)
                            continue
                result = manager.run_command(target["object"], "GET_POOLS_STATUS_COMMAND")
                if not result["stdout"]:
                    show_get_pools_error(target, result)
                    continue
                objects = output_to_pool_objects(hostname=target["object"].hostname, command_output=result["stdout"])
                for iis_pool in objects:
                    show_formatted_pool_attributes(iis_pool, "STOP_APP_POOL_COMMAND", result["status_code"])
            current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            print(f"{Fore.LIGHTWHITE_EX}[ FINISHED AT {current_time} ]{Style.RESET_ALL}")

        elif menu_choice == "4":
            current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            print(f"{Fore.LIGHTWHITE_EX}[ STARTED AT {current_time} ]{Style.RESET_ALL}")
            show_output_header()
            for target in manager.target_objects.values():
                result = manager.run_command(target["object"], "GET_POOLS_STATUS_COMMAND")
                if not result["stdout"] or result["status_code"] != 0:
                    show_get_pools_error(target, result)
                    continue
                objects = output_to_pool_objects(hostname=target["object"].hostname, command_output=result["stdout"])
                for iis_pool in objects:
                    show_formatted_pool_attributes(iis_pool, "GET_POOLS_STATUS_COMMAND", result["status_code"])
            current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            print(f"{Fore.LIGHTWHITE_EX}[ FINISHED AT {current_time} ]{Style.RESET_ALL}")
        elif menu_choice == "0":
            sys.exit(0)
        else:
            main()


def output_to_pool_objects(hostname, command_output):
    """converts pools attributes command output to object, then add they to a list"""
    objects = []
    splited_output = command_output.replace("\r", "").lstrip("\n").rstrip("\n").split("\n\n")
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


def show_formatted_pool_attributes(iis_pool, command, status_code):
    if status_code == 0:
        result_message = f"Successfull with status code: {status_code}"
    else:
        result_message = f"Failed with status code: {status_code}"
    print(f"{iis_pool.hostname.ljust(20, ' ')}{iis_pool.name.ljust(20, ' ')}{iis_pool.state.ljust(20, ' ')}"
          f"{iis_pool.auto_start.ljust(20, ' ')}{iis_pool.start_mode.ljust(20, ' ')}{command.ljust(30, ' ')}{result_message}")


def show_iisreset_output(hostname, command_output):
    print(f"{hostname.ljust(20, ' ')}{'Successfuly reseted'.ljust(80, ' ')}")


def show_get_pools_error(target, result):
    print(f"{target['object'].ip_address.ljust(100, ' ')}{'GET_POOLS_STATUS_COMMAND'.ljust(30, ' ')}"
          f"Failed with status code: {result['status_code']}")


def show_output_header():
    print(f"{'SERVER'.ljust(20, ' ')}{'POOL NAME'.ljust(20, ' ')}{'POOL STATE'.ljust(20, ' ')}"
          f"{'POOL AUTO START'.ljust(20, ' ')}{'POOL START MODE'.ljust(20, ' ')}{'COMMAND'.ljust(30, ' ')}"
          f"{'MESSAGE'.ljust(40)}\n{160 * '-'}")


if __name__ == "__main__":
    main()

