try:
    import os
    import uuid
    from FileOperations import FileOperations
except ImportError as error:
    print(f"Import Error..... Script Stopped.. {error}")
    exit(1)


class IndiaSpeakUp:

    def __init__(self, logFolder=None, logger=None, guid=None):

        self.start_time = time.time()
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
            "useCaseName": "GI-Ethics",
        }
        self.file_obj = FileOperations(__file__, self.logger, self.guid)

