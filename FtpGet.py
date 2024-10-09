import os
import shutil
import paramiko
from datetime import datetime
import logging

logging.basicConfig(level=logging.DEBUG,
                    format="%(asctime)s - %(levelname)s - %(message)s",
                    handlers=[logging.StreamHandler(),
                    logging.FileHandler("FtpGet.log")])
logging.info("FtpGet Started \n")

#This string of spaces is for formatting log reports nicely
log_tab ="                                 "

#Get FTP Server Details from System Variables
sftp_username = os.getenv('FtpUserName')
sftp_password = os.getenv('FtpUserPass')
sftp_server = os.getenv('FtpHost')
remote_folder = 'Outbound'
today_date = datetime.now().strftime('%Y%m%d')
if sftp_username == None or sftp_password == None or sftp_server == None:
    logging.error("Required Environmental Variables not found, \n"+log_tab
                +"this script requires three environmental\n"+log_tab+
                "variables to work.\n"+log_tab+" 1) FtpUserName\n"+log_tab+
                " 2) FtpUserPass\n"+log_tab+" 3) FtpHost\n")
    logging.warning("Terminating Script Early")
    exit()

#Set the folder to save files to when downloaded from FTP
local_folder = "c:\\FTP-Down"
try:
    if not os.path.exists(local_folder):
        os.makedirs(local_folder)
        logging.info(f"Created target folder {local_folder}")
    else:
        logging.info(f"Setting target folder to {local_folder}")
except Exception as e:
    logging.error(f"Error: {str(e)}")

#Set the server folder to move the final files to
server_folder = "\\\\server19\\db\\Paycom Reports\\Paycom Data"
if os.path.exists(server_folder) and os.access(server_folder, os.W_OK):
    logging.info(f"Connected to server folder:\n"+log_tab+server_folder)
else:
    logging.error("Cannot connect to server folder:\n "+log_tab+server_folder+
                "\n"+log_tab+" Please verify folder exists and this profile\n"
                +log_tab+" has access")
    exit()

#Dictionary showing expected file name beginnings and what the file name 
#should be change to before moving it to paycom data
file_rename_matrix = {
    "Luci Allocations Report": "Luci Allocations Report.xlsx",
    "Project_OH Time by Employee_Department_Location_LUCI": "Luci Hours 2024.xlsx",
    "Project_OH Time by Employee_Department_Location": "2024 Labor Hours.xlsx",
    "Punch Record": "Punches Current Quarter.xlsx",
    "RSI Allocations Report v2": "RSI Allocations Report.xlsx",
    "RSI_Job_Totals_Active": "RSI_Job_Totals_Active.xlsx"
    }

def download_files():
    try:
        #Create an SSH Transport Client
        transport = paramiko.Transport((sftp_server, 22))
        transport.connect(username=sftp_username, password=sftp_password)
        #Create the FTP session
        sftp = paramiko.SFTPClient.from_transport(transport)
        logging.info(f"Opened SFTP Connection: {sftp_server}")
        sftp.chdir(remote_folder)
        for filename in sftp.listdir():
            if filename.startswith(today_date):
                local_file_path = os.path.join(local_folder, filename)

                sftp.get(filename, local_file_path)
                print(f'Downloaded: {filename}')
            
        print('Closing SFTP Connection')
        sftp.close()
        transport.close()

    except Exception as e:
        print(f'An error occurred: {e}')

def strip_date():
    try:
        os.chdir(local_folder)
        for filename in os.listdir(local_folder):
            if filename.startswith(today_date):
                new_filename = filename[15:]
                old_filepath = os.path.join(local_folder, filename)
                new_filepath = os.path.join(local_folder, new_filename)
                os.rename(old_filepath, new_filepath)
                print(f"Renamed {filename} to {new_filename}")
        print("Date info stripped from all filenames")

    except Exception as e:
        print(f'An error occurred: {e}')
        
def rename_files():
    try:
        for filename in os.listdir(local_folder):
            for prefix, new_name in file_rename_matrix.items():
                if filename.startswith(prefix):
                    new_filename = new_name
                    break
            else:
                continue
            old_filepath = os.path.join(local_folder, filename)
            new_filepath = os.path.join(local_folder, new_filename)
            os.rename(old_filepath, new_filepath)
            print(f"Renamed {filename} to {new_filename}")
        print('Rename Function Complete')

    except Exception as e:
        print(f"Error: {str(e)}")

def move_files():
    try:
        for filename in os.listdir(local_folder):
            local_name = os.path.join(local_folder, filename)
            remote_name = os.path.join(server_folder, filename)
            shutil.move(local_name, remote_name)
            print(f"Moved {filename} to {server_folder}")

    except Exception as e:
        print(f"Error {str(e)}")

def main():
    download_files()

#    strip_date()
#    rename_files()
#    move_files()
    return

if __name__ == "__main__":
    main()
logging.info("FtpGet Complete \n\n")