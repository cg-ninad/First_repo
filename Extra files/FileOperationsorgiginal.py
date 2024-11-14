try:
    import os
    import uuid
    import glob
    import shutil
    import requests
    import datetime
    import pandas as pd
    from logger_format import setup_logging
    from Config import SNOW, SC_PSWD, SC_USER
except ImportError as import_error:
    print(f"Import Error... Script stopped {import_error}")
    exit(1)


class FileOperations:

    def __init__(self, logFolder=None, logger=None, guid=None):
        # Creating logger object
        if guid:
            self.guid = guid
        else:
            self.guid = uuid.uuid4()
        if logFolder:
            self.logger = logger
        else:
            self.logger = setup_logging(__file__)
        self.extra_dict_common = {
            "JobId": self.guid,
            "logType": "User",
            "exec_category": "Python_Script",
            "Category": "Self Heal",
            "businessCategory": "GI",
            "useCaseName": "Data Retention",
        }

    def get_file_name(self, shared_path):
        """
        Get file name uploaded on shared folder.
        File name to process data.
        :return:
        return filename with it's path
        """
        try:
            if os.path.exists(shared_path):
                files = glob.glob(shared_path + '/*.xlsx')
                if len(files) == 0:
                    self.logger.error(f"No *.xlsx file present in {shared_path} shared folder, Please upload file to process data",
                                      extra=self.extra_dict_common)
                    msg = f"""No *.xlsx file present in {shared_path} shared folder, please upload respective file 
                            in given path to process data. 
                            <mark>"<u>SpeakUp</u>" word must be in file name for SpeakUP Alerts file.</mark>"""
                    return [False, msg]
                return [True, files[0]]
            else:
                self.logger.error(f"get_file_name();{shared_path} File Path not exists", extra=self.extra_dict_common)
                return [False, "shared path does not exist"]
        except Exception as err:
            self.logger.exception(f"get_file_name(); Failed to get file from shared folder {err}",
                                  extra=self.extra_dict_common)
            return [False, err]

    def move_files_archive(self, source_file, destination_dir):
        """
        move file from source_dir to destination_dir to avoid ambiguity
        :param source_file:
        :param destination_dir:
        :return:
        """
        try:

            real_dst = os.path.join(destination_dir, os.path.basename(source_file))
            if not os.path.exists(real_dst):
                result = shutil.move(source_file, destination_dir)
            else:
                fname = f"{datetime.datetime.now().strftime('%d-%m-%Y')}_{os.path.basename(source_file)}"
                dst = os.path.join(destination_dir, fname)
                result = shutil.move(source_file, dst)

            self.logger.info(
                f"move_files_archive(); file moved {result}", extra=self.extra_dict_common
            )
            return [True, "File Removed"]
        except Exception as error:
            self.logger.exception(f"move_files_archive(); Failed to move file; {error}")
            return [False, error]

    def read_file(self, filepath):
        """
        Read file from shared directory and return dataframe as a output
        :param filepath:
        :return:
        """
        try:
            if os.path.exists(filepath):
                if '.xlsx' in filepath.__str__():
                    try:
                        df = pd.read_excel(filepath)
                    except Exception:
                        df = pd.read_excel(filepath)
                    return [df]
                else:
                    self.logger.exception("read_file(); Other than '.xlsx' file uploaded. Failed to process",
                                          extra=self.extra_dict_common)
                    er = "Invalid file format found in the shared folder, only .xlsx file will be accepted. "
                    return [False, er]
            else:
                self.logger.error(
                    "read_file(); No file present in folder; {}".format(filepath),
                    extra=self.extra_dict_common,
                )
                f_error = f"File path/ file does not exist for given path {filepath}"
                return [False, f_error]
        except Exception as error:
            self.logger.exception(f"read_file(); Failed to read file; {error}")
            return [False, error]

    def delete_file(self, filename):
        """
        delete file from given path
        :param filename:
        :return:
        """
        try:
            if os.path.exists(filename):
                os.unlink(filename)
        except Exception as error:
            self.logger.exception(f"delete_file(); Failed to delete file; {error}", extra=self.extra_dict_common)

    def files_cleanup(self):
        try:
            files = glob.glob("*.xlsx")
            for file in files:
                os.remove(file)
            return True
        except:
            pass

    def get_email_id(self, corp_id=None):

        try:
            header = {'Content-Type': 'application/json', 'Accept': 'application/json'}
            if corp_id is None:
                return None
            corp_id = int(corp_id)

            url = f"{SNOW}sysparm_fields=email&sysparm_limit=1&u_global_id={corp_id}"
            response = requests.get(url,auth=(SC_USER, SC_PSWD),headers=header)
            sys_dict = response.json()
            mail = sys_dict.get("result", [])[0].get("email", "")
            return mail
        except Exception as e:
            self.logger.exception(f"get_email_id(); Corp Id:{corp_id}, Error: {e}", extra=self.extra_dict_common)
            pass
