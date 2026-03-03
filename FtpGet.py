import os
import sys
import shutil
import paramiko
from paramiko.ssh_exception import SSHException, NoValidConnectionsError
from datetime import datetime
import logging
import textwrap
import smtplib
import json
from email.message import EmailMessage
from pathlib import Path


class IndentedFormatter(logging.Formatter):
    def format(self, record):
        # Set the width for the message wrap, adjusting for
        # the lenght of other log fields
        width = 100
        indent = " " * 31

        original_msg = super().format(record)
        wrapped_msg = textwrap.fill(original_msg, width=width, subsequent_indent=indent)
        return wrapped_msg


#configuration of logging streams
ftpget_logger = logging.getLogger('internal_logger')
ftpget_logger.setLevel(logging.DEBUG)

stream_handler = logging.StreamHandler()
file_hanlder = logging.FileHandler('FtpGet.log')

ftpget_formatter = IndentedFormatter('%(asctime)s - %(levelname)s - %(message)s')
stream_handler.setFormatter(ftpget_formatter)
file_hanlder.setFormatter(ftpget_formatter)

ftpget_logger.addHandler(stream_handler)
ftpget_logger.addHandler(file_hanlder)

ftpget_logger.info("FtpGet Started \n")


# --- Email Configuration ---

EMAIL_FROM = os.getenv('EMAIL_FROM', 'paycomftpbot@ravenswoodstudio.com')
SMTP_SERVER = os.getenv('SMTP_SERVER')
SMTP_PORT = int(os.getenv('SMTP_PORT', '25'))

if SMTP_SERVER is None:
    ftpget_logger.warning("***************************\n"
                          "  SMTP_SERVER var missing  \n"
                          " Proceeding without email  \n"
                          "***************************\n")

# Load recipient list from recipients.json in the same directory as this script
_recipients_path = Path(__file__).resolve().parent / "recipients.json"
try:
    with open(_recipients_path, 'r') as f:
        _recipients_data = json.load(f)
    EMAIL_RECIPIENTS = _recipients_data.get("recipients", [])
    if not EMAIL_RECIPIENTS:
        ftpget_logger.warning("recipients.json loaded but recipient list is empty.")
except FileNotFoundError:
    ftpget_logger.warning(f"recipients.json not found at {_recipients_path}. No emails will be sent.")
    EMAIL_RECIPIENTS = []
except Exception as e:
    ftpget_logger.warning(f"Failed to load recipients.json: {e}. No emails will be sent.")
    EMAIL_RECIPIENTS = []


def send_email(subject: str, body: str) -> None:
    """
    Send a notification email via anonymous internal SMTP relay.

    Required env vars:
        SMTP_SERVER : hostname or IP of the internal SMTP relay
        SMTP_PORT   : port to use (defaults to 25)
        EMAIL_FROM  : sender address (defaults to paycomftpbot@ravenswoodstudio.com)

    Recipients are loaded from recipients.json in the script directory.
    """
    if not SMTP_SERVER:
        ftpget_logger.warning("SMTP_SERVER is not set; email not sent.")
        return
    if not EMAIL_RECIPIENTS:
        ftpget_logger.warning("No recipients configured; email not sent.")
        return

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = EMAIL_FROM
    msg['To'] = ', '.join(EMAIL_RECIPIENTS)
    msg.set_content(body)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.sendmail(EMAIL_FROM, EMAIL_RECIPIENTS, msg.as_string())
        ftpget_logger.info(f"Email sent: {subject}")
    except Exception as e:
        ftpget_logger.warning(f"Failed to send email: {e}")


def get_env_var(var_name: str, is_required: bool = True) -> str:
    value = os.getenv(var_name)
    if value is None:
        if is_required:
            message = f"{var_name} is missing. Please set it in the environment variables"
            ftpget_logger.error(message)
            send_email(
                subject="FtpGet - Missing Environment Variable",
                body=message
            )
            exit()
        return ""
    return value


#Get FTP Server Details from System Variables
sftp_username = get_env_var('FtpUserName')
sftp_password = get_env_var('FtpUserPass')
sftp_server = get_env_var('FtpHost')

remote_folder = 'Outbound'
today_date = datetime.now().strftime('%Y%m%d')
 
#Set the folder to save files to when downloaded from FTP
local_folder = Path(__file__).resolve().parent / "FTP-Down"
try:
    if not os.path.exists(local_folder):
        os.makedirs(local_folder)
        ftpget_logger.info(f"Created target folder {local_folder}")
    else:
        for file_name in os.listdir(local_folder):
            file_path = os.path.join(local_folder, file_name)
            if os.path.isfile(file_path):
                os.remove(file_path)
                ftpget_logger.info(f"Deleted: {file_path} from target folder")
        ftpget_logger.info(f"Setting target folder to {local_folder}")
except Exception as e:
    ftpget_logger.error(f"Error: {str(e)}")

#Set the server folder to move the final files to
server_folder = Path(r'\\server19\db\Paycom Reports')
if os.path.exists(server_folder) and os.access(server_folder, os.W_OK):
    ftpget_logger.info(f"Connected to server folder: {server_folder}")
else:
    ftpget_logger.error(f"Cannot connect to server folder: {server_folder} "
                        "Please verify folder exists and this profile has access")
    send_email(
        subject="FtpGet - Initialization Failed",
        body=f"Automation failed to initialize. Cannot access Server Folder: {server_folder}\n"
             "Please verify the folder exists and the service account has write access."
    )
    exit()
unexpected_subfolder = server_folder / "Unexpected-Reports"


#Dictionary where each primary key is a string giving the expected file name
#prefix, each primary entry contains two secondary keys 'newname' contains 
#the full filename it should be renamed to, and 'folder' gives us a path
#to where the file should be moved to within the path initialized in 
#server_folder
file_rename_matrix = {
    "Luci Allocations Report":{
        'newname':"Luci Allocations Report.xlsx",
        'folder':"Paycom Data"
    },
    "Project_OH Time by Employee_Department_Location_LUCI":{
        'newname':"Luci Hours 2025.xlsx",
        'folder':"Paycom Data"
    },
    "Project_OH Time by Employee_Department_Location_RSI":{
        'newname':"Current Labor Hours.xlsx",
        'folder':"Paycom Data"
    },
    "RSI Allocations Report v2":{
        'newname':"RSI Allocations Report.xlsx",
        'folder':"Paycom Data"
    },
    "RSI_Job_Totals_Active":{
        'newname':"RSI_Job_Totals_Active.xlsx",
        'folder':"Paycom Data"
    },
    "Data_RSI_Labor Allocation Summary":{
        'newname':"THSA_RSI_YTD.csv",
        'folder':r"Ravenswood Studio\Dashboard\Data\THSA"
    },
    "Data_RSI_Punch Record":{
        'newname':"PR_RSI_QTD.csv",
        'folder':r"Ravenswood Studio\Dashboard\Data\PR"
    },
    "Data_RSI_Time Detail Report":{
        'newname':"TDR_RSI_YTD.csv",
        'folder':r"Ravenswood Studio\Dashboard\Data\TDR"
    }
}


def download_files(expected_count) -> int:
    dl_count = 0
    transport: paramiko.Transport | None = None
    sftp: paramiko.SFTPClient | None = None
    try:
        #Create an SSH Transport Client
        transport = paramiko.Transport((sftp_server, 22))
        transport.connect(username=sftp_username, password=sftp_password)
        #Create the FTP session
        sftp = paramiko.SFTPClient.from_transport(transport)
        if sftp is None:
            raise SSHException("Failed to create SFTP session from transport")
        ftpget_logger.info(f"Opened SFTP Connection: {sftp_server}")
        sftp.chdir(remote_folder)
        for filename in sftp.listdir():
            if filename.startswith(today_date):
                local_file_path = os.path.join(local_folder, filename)
                sftp.get(filename, local_file_path)
                ftpget_logger.info(f'Downloaded: {filename}')
                dl_count += 1
      
    except (SSHException, NoValidConnectionsError) as e:
        ftpget_logger.error(f"Connection Failed: {e}")
        send_email(
            subject="FtpGet - SFTP Connection Failed",
            body=f"Could not connect to SFTP server ({sftp_server}), operation aborted.\n\nError: {e}"
        )
        exit()
    except Exception as e:
        ftpget_logger.error(f'An error occurred: {e}')
        send_email(
            subject="FtpGet - Unknown Download Error",
            body=f"An unexpected error occurred during the download operation and the process was aborted.\n\nError: {e}"
        )
        exit()
    finally:
        ftpget_logger.info('Closing SFTP Connection')
        if sftp is not None:
            sftp.close()
        if transport is not None:
            transport.close()
    return dl_count - expected_count


def strip_date() -> None:
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
        ftpget_logger.error("Aborting")
        send_email(
            subject="FtpGet - File Rename Error",
            body=f"Something went wrong stripping date prefixes from filenames. Process aborted.\n\nError: {e}"
        )
        exit()      


def rename_files() -> list[str]:
    file_list = [file['newname'] for file in file_rename_matrix.values()]
    try:
        for filename in os.listdir(local_folder):
            for prefix in file_rename_matrix.keys():
                if filename.startswith(prefix):
                    new_filename = file_rename_matrix[prefix]["newname"]
                    break
            else:
                continue
            old_filepath = os.path.join(local_folder, filename)
            new_filepath = os.path.join(local_folder, new_filename)
            os.rename(old_filepath, new_filepath)
            ftpget_logger.info(f"Renamed {filename} to {new_filename}")
            file_list.remove(new_filename)
        ftpget_logger.info('Rename Function Complete')
    
    except Exception as e:
        ftpget_logger.error(f"Error: {str(e)}")
        ftpget_logger.error("Aborting")
        send_email(
            subject="FtpGet - File Rename Error",
            body=f"Something went wrong renaming files. Process aborted.\n\nError: {e}"
        )
        exit()
    return file_list


def move_files(missing_files) -> None:
    try:
        move_matrix = {details["newname"]: details["folder"] for details in file_rename_matrix.values()}
        for filename in missing_files:
            del move_matrix[filename]
        for filename, folder in move_matrix.items():
            try:
                local_name = os.path.join(local_folder, filename)
                remote_name = os.path.join(server_folder, folder, filename)
                shutil.move(local_name, remote_name)
                ftpget_logger.info(f"Moved {filename} to {server_folder}\\{folder}")

            except Exception as e:
                ftpget_logger.error(f"Error {str(e)}")
                ftpget_logger.error(f"Could not move {filename}")
                send_email(
                    subject=f"FtpGet - File Move Error: {filename}",
                    body=f"There was a problem moving {filename} to {server_folder}\\{folder}.\n\nError: {e}"
                )

    except Exception as e:
        ftpget_logger.error(f"Error {str(e)}")
        ftpget_logger.error("Aborting, files may not be moved to server.")
        send_email(
            subject="FtpGet - Critical File Move Error",
            body=f"There was a problem moving files to {server_folder}. Process aborted.\n\nError: {e}"
        )
        exit()


def move_extra_files() -> None:
    #Function to move files left over after the rename and move to server but 
    #were downloaded because they had the correct datestamp for todays download.
    #They are not yet defined in the File Rename matrix so we move them to a
    #subdirectory in the server folder for inspection. 
    try:
        if not os.path.exists(unexpected_subfolder):
            os.makedirs(unexpected_subfolder)
        for filename in os.listdir(local_folder):
            local_name = os.path.join(local_folder, filename)
            remote_name = os.path.join(unexpected_subfolder, filename)
            shutil.move(local_name, remote_name)
            ftpget_logger.info(f"Moved {filename} to {unexpected_subfolder}")
    except Exception as e:
        ftpget_logger.error(f"Error {str(e)}")
        ftpget_logger.error("Aborting, unexpected files may not be moved to server.")
        send_email(
            subject="FtpGet - Unexpected Files Move Error",
            body=f"There was a problem moving unexpected files to {unexpected_subfolder}.\n\nError: {e}"
        )


def main() -> None:
    expected_count = len(file_rename_matrix)
    dl_dif = download_files(expected_count)

    if dl_dif > 0:
        message = f"There were {dl_dif} more file(s) downloaded than expected."
        ftpget_logger.info(message)

    elif dl_dif < 0:
        message = f"{abs(dl_dif)} of the expected {expected_count} file(s) were missing from the FTP download."
        ftpget_logger.info(message)

    strip_date()
    missing_files = rename_files()
    move_files(missing_files)

    # Build a summary email if there were any exceptions
    exceptions = []

    if dl_dif > 0:
        extra_files = os.listdir(local_folder)
        extra_list = "\n".join(f"  - {f}" for f in extra_files)
        exceptions.append(
            f"The following unexpected file(s) were downloaded and moved to {unexpected_subfolder}:\n{extra_list}"
        )
        move_extra_files()

    if len(missing_files) > 0:
        missing_list = "\n".join(f"  - {f}" for f in missing_files)
        exceptions.append(
            f"The following expected file(s) were not found and were not updated:\n{missing_list}"
        )
        for file in missing_files:
            ftpget_logger.warning(file)

    if exceptions:
        body = "The Paycom FTP automation completed with the following exceptions:\n\n"
        body += "\n\n".join(exceptions)
        send_email(
            subject="FtpGet - Update Complete With Exceptions",
            body=body
        )
        ftpget_logger.info("FtpGet Complete with exceptions\n\n")
    else:
        send_email(
            subject="FtpGet - Update Complete",
            body=f"Paycom data updated successfully. All {expected_count} files downloaded and moved to the server."
        )
        ftpget_logger.info("FtpGet Complete \n\n")

    return


if __name__ == "__main__":
    main()
