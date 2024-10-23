import os
import shutil
import paramiko
from datetime import datetime
import logging

#configuration of logging streams
ftpget_logger = logging.getLogger('internal_logger')
ftpget_logger.setLevel(logging.DEBUG)

stream_handler = logging.StreamHandler()
file_hanlder = logging.FileHandler('FtpGet.log')

ftpget_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
stream_handler.setFormatter(ftpget_formatter)
file_hanlder.setFormatter(ftpget_formatter)

ftpget_logger.addHandler(stream_handler)
ftpget_logger.addHandler(file_hanlder)

external_logger = logging.getLogger('external_module_logger')
external_logger.setLevel(logging.ERROR)
external_handler = logging.FileHandler('modules.log')

external_handler.setFormatter(ftpget_formatter)
external_logger.addHandler(external_handler)
paramiko.logger = external_logger

# Old Logger - Remove once new Logger is fully functional
#logging.basicConfig(level=logging.DEBUG,
#                    format="%(asctime)s - %(levelname)s - %(message)s",
#                    handlers=[logging.StreamHandler(),
#                    logging.FileHandler("FtpGet.log")])
ftpget_logger.info("FtpGet Started \n")

#This string of spaces is for formatting log reports nicely
log_tab ="                                 "

#Get FTP Server Details from System Variables
sftp_username = os.getenv('FtpUserName')
sftp_password = os.getenv('FtpUserPass')
sftp_server = os.getenv('FtpHost')
remote_folder = 'Outbound'
today_date = datetime.now().strftime('%Y%m%d')
if sftp_username == None or sftp_password == None or sftp_server == None:
    ftpget_logger.error("Required Environmental Variables not found, \n"+log_tab
                +"this script requires three environmental\n"+log_tab+
                "variables to work.\n"+log_tab+" 1) FtpUserName\n"+log_tab+
                " 2) FtpUserPass\n"+log_tab+" 3) FtpHost\n")
    ftpget_logger.warning("Terminating Script Early")
    exit()

#Set the folder to save files to when downloaded from FTP
local_folder = "c:\\FTP-Down"
try:
    if not os.path.exists(local_folder):
        os.makedirs(local_folder)
        ftpget_logger.info(f"Created target folder {local_folder}")
    else:
        ftpget_logger.info(f"Setting target folder to {local_folder}")
except Exception as e:
    ftpget_logger.error(f"Error: {str(e)}")

#Set the server folder to move the final files to
server_folder = "\\\\server19\\db\\Paycom Reports\\Test-Paycom Data"
if os.path.exists(server_folder) and os.access(server_folder, os.W_OK):
    ftpget_logger.info(f"Connected to server folder:\n"+log_tab+server_folder)
else:
    ftpget_logger.error("Cannot connect to server folder:\n "+log_tab+server_folder+
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
        ftpget_logger.info(f"Opened SFTP Connection: {sftp_server}")
        sftp.chdir(remote_folder)
        for filename in sftp.listdir():
            if filename.startswith(today_date):
                local_file_path = os.path.join(local_folder, filename)

                sftp.get(filename, local_file_path)
                ftpget_logger.info(f'Downloaded: {filename}')
            
        ftpget_logger.info('Closing SFTP Connection')
        sftp.close()
        transport.close()

    except Exception as e:
        ftpget_logger.error(f'An error occurred: {e}')

def strip_date():
    try:
        os.chdir(local_folder)
        for filename in os.listdir(local_folder):
            if filename.startswith(today_date):
                new_filename = filename[15:]
                old_filepath = os.path.join(local_folder, filename)
                new_filepath = os.path.join(local_folder, new_filename)
                os.rename(old_filepath, new_filepath)
                ftpget_logger.info(f"Renamed {filename} to {new_filename}")
        ftpget_logger.info("Date info stripped from all filenames")

    except Exception as e:
        ftpget_logger.error(f'An error occurred: {e}')
        
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
            ftpget_logger.info(f"Renamed {filename} to {new_filename}")
        ftpget_logger.info('Rename Function Complete')

    except Exception as e:
        ftpget_logger.error(f"Error: {str(e)}")

def move_files():
    try:
        for filename in os.listdir(local_folder):
            local_name = os.path.join(local_folder, filename)
            remote_name = os.path.join(server_folder, filename)
            shutil.move(local_name, remote_name)
            ftpget_logger.info(f"Moved {filename} to {server_folder}")

    except Exception as e:
        ftpget_logger.error(f"Error {str(e)}")

def main():
    download_files()
    strip_date()
    rename_files()
    move_files()
    ftpget_logger.info("FtpGet Complete \n\n")
    return

if __name__ == "__main__":
    main()