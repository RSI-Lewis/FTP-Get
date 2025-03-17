import os
import shutil
import paramiko
from paramiko.ssh_exception import SSHException, NoValidConnectionsError
from datetime import datetime
import logging
import textwrap
from slack_sdk import WebClient
from pathlib import Path
import time


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

#initiate SlackBot connection
slack_token=os.getenv('slack_auth')
if slack_token == None:
    ftpget_logger.warning("***************************\n"+
                        " Slack Auth Token Missing\n"+
                        "Proceeding without Slackbot\n"+
                        "***************************\n")
else:
    client = WebClient(token=slack_token)


def post_to_slack(message, channel="paycom-automation", username="Bot User") -> None:
    """
    Function to post a mesasge to SlackBot
    
    Required parameter:
        message (str): Contains the message to be posted.
    
    Optional parameters
        channel (str): channel to post into defautls to paycom-automation
        username (str): username to post from, defaults to "Bot User"
    """

    if slack_token:
        try:
            client.chat_postMessage(channel=channel, text=message, username=username)
        except Exception as e:
            ftpget_logger.warning(f"Failed to send message to Slack: {e}")
    else:
        ftpget_logger.warning("Slack token is missing; message not sent.")


def get_env_var(var_name, is_required=True) -> str:
    value = os.getenv(var_name)
    if is_required and value is None:
        message = f"{var_name} is missing. Please set it in the environment variables"
        ftpget_logger.error(message)
        post_to_slack(message)
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
        ftpget_logger.info(f"Setting target folder to {local_folder}")
except Exception as e:
    ftpget_logger.error(f"Error: {str(e)}")

#Set the server folder to move the final files to
server_folder = Path(r'\\server19\db\Paycom Reports\Paycom Data')
if os.path.exists(server_folder) and os.access(server_folder, os.W_OK):
    ftpget_logger.info(f"Connected to server folder: {server_folder}")
else:
    ftpget_logger.error(f"Cannot connect to server folder: {server_folder}"+
                "Please verify folder exists and this profile has access")
    post_to_slack("Automation failed to initialize. Cannot access Server Folder")
    exit()
unexpected_subfolder = server_folder / "Unexpected-Reports"


#Dictionary showing expected file name beginnings and what the file name 
#should be change to before moving it to paycom data
file_rename_matrix = {
    "Luci Allocations Report": "Luci Allocations Report.xlsx",
    "Project_OH Time by Employee_Department_Location_LUCI": "Luci Hours 2025.xlsx",
    "Project_OH Time by Employee_Department_Location": "2025 Labor Hours.xlsx",
    "Punch Record": "Punches Current Quarter.xlsx",
    "RSI Allocations Report v2": "RSI Allocations Report.xlsx",
    "RSI_Job_Totals_Active": "RSI_Job_Totals_Active.xlsx"
    }


def download_files(expected_count) -> int:
    dl_count = 0
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
                dl_count += 1
      
    except (SSHException, NoValidConnectionsError) as e:
        ftpget_logger.error(f"Connection Failed: {e}")
        post_to_slack("Could not connect to SFTP server, operation aborted")
        exit()
    except Exception as e:
        ftpget_logger.error(f'An error occurred: {e}')
        post_to_slack("Unknown Exception in download operation, aborted")
        exit()
    finally:
        ftpget_logger.info('Closing SFTP Connection')
        sftp.close()
        transport.close()
    return dl_count-expected_count


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
        post_to_slack("Something went wrong renaming files. Aborted")
        exit()      


def rename_files() -> list[str]:
    file_list = list(file_rename_matrix.values())
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
            file_list.remove(new_filename)
        ftpget_logger.info('Rename Function Complete')
    
    except Exception as e:
        ftpget_logger.error(f"Error: {str(e)}")
        ftpget_logger.error("Aborting")
        post_to_slack("Something went wrong renaming files. Aborted")
        exit()
    return file_list


def move_files() -> None:
    time.sleep(1)
    try:
        file_list = list(file_rename_matrix.values())
        for filename in file_list:
            local_name = os.path.join(local_folder, filename)
            remote_name = os.path.join(server_folder, filename)
            shutil.move(local_name, remote_name)
            ftpget_logger.info(f"Moved {filename} to {server_folder}")

    except Exception as e:
        ftpget_logger.error(f"Error {str(e)}")
        ftpget_logger.error("Aborting,files may not be moved to server.")
        post_to_slack(f"There was a problem moving files to {server_folder}")
    #TODO find out why this exception does not make a post to the slack
        exit()


def move_extra_files() -> None:
    #Function to move files left over after the rename and move to server but 
    #were downloaded because they had the correct datastamp for todays download
    #They are not yet defined in the File Rename matrix so we move them to a
    #subdirectory in the server folder for inspection 
    
    try:
        if not os.path.exists(unexpected_subfolder):
            os.makedirs(unexpected_subfolder)
        for filename in os.listdir(local_folder):
            local_name = os.path.join(local_folder, filename)
            remote_name = os.path.join(unexpected_subfolder, filename)
            shutil.move(local_name, remote_name)
            ftpget_logger.info(f"Moved {filename} to {unexpected_subfolder}")
            post_to_slack(text=f"Unexpected files moved to {unexpected_subfolder}")
    except Exception as e:
        ftpget_logger.error(f"Error {str(e)}")
        ftpget_logger.error("Aborting,unexpected files may not be moved to server.")
        post_to_slack(f"There was a problem moving unexpected files to {unexpected_subfolder}")


def main() -> None:
    expected_count = len(file_rename_matrix)
    dl_dif = download_files(expected_count)
    if dl_dif == 0:
        pass
    elif dl_dif > 0:
        message = f"There was {dl_dif} more file(s) downloaded than expected"
        ftpget_logger.info(message)
        post_to_slack(message)
    elif dl_dif < 0:
        message = f"{abs(dl_dif)} file(s) of expected {expected_count} were missing"
        ftpget_logger.info(message)
        post_to_slack(message)

    strip_date()
    missing_files = rename_files()
    move_files()
    
    if dl_dif > 0:
        message = "These extra files were not in my database:"
        ftpget_logger.warning(message)
        post_to_slack(message)
        missed_files = os.listdir(local_folder)
        for file in missed_files:
            ftpget_logger.warning(file)
            post_to_slack(file)
        move_extra_files()
        ftpget_logger.info("FtpGet Complete with exceptions\n\n")
        post_to_slack("Update complete with above exceptions")
    elif dl_dif < 0:
        message = "These file(s) were not updated:"
        ftpget_logger.warning(message)
        post_to_slack(message)
        for file in missing_files:
            ftpget_logger.warning(file)
            post_to_slack(file)
        ftpget_logger.info("FtpGet Complete with exceptions\n\n")
        post_to_slack("Update complete with above exceptions")

    ftpget_logger.info("FtpGet Complete \n\n")
    post_to_slack(f"Paycom data updated, all {expected_count} files complete.")
    return


if __name__ == "__main__":
    main()