try:
    import os
    import uuid
    import datetime
    from outlook_exch_lib import ExchangeMail
    from logger_format import setup_logging
    from Config import MAIL_USER, MAIL_ACCOUNT, MAIL_PASS, SPEAK_SUB, CC, BUSINESS, BCC, SUPPORT
    from htmls import HEAD, SIGNATURE_SUPPORT, HEADER, CONTENT
except Exception as error:
    print(f"Import Error.. {error}")
    exit(1)


class MailList:

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
            "useCaseName": "GI-Ethics",
        }
        self.outlookobj = ExchangeMail(MAIL_ACCOUNT, MAIL_USER, MAIL_PASS, __file__, self.logger, self.guid)

    def file_missing_speakup(self, errors=''):

        subject = f"Data Retention timeline | Report Status | BOT process failed"
        today = datetime.datetime.now().strftime('%d-%m-%Y')
        mail_body = f"""<p>
                Hi Team, <br>
                This is to inform you that the BOT process failed to process the SpeakUp report or 
                to send emails to the recipients as of {today}. <br>
                <b><u>Error:</u></b>
                <p style="font-size:12.5px;">
                    {errors}
                </p>
                <br><br>
                Thanks,<br>
                AutoBot</p>                
        """
        return [subject, mail_body]

    def file_missing_declare1(self, errors=''):
        subject = f"Data Retention timeline | Report Status | BOT process failed"
        today = datetime.datetime.now().strftime('%d-%m-%Y')
        mail_body = f"""<p>
                Hi Team, <br>
                This is to inform you that the BOT process failed to process the Declare report or to send emails to 
                the recipients as of  {today}. <br>
                <b><u>Error:</u></b>
                <p style="font-size:12.5px;">
                    {errors}
                </p>
                <br><br>
                Thanks,<br>
                AutoBot</p>                
        """
        return [subject, mail_body]

    def send_mail(self, subject=None, mailbody=None, to=BUSINESS, cc=None, bcc=BCC, attachments=None):

        try:
            msg_flg = self.outlookobj.SendMail(
                to_recipients=to,
                cc_recipients=cc,
                bcc_recipients=bcc,
                mail_subject=subject,
                mail_body=mailbody,
                is_htmlbody=True,
                mail_attachments=attachments,
            )
            self.logger.info(f"Mail Status: {msg_flg}", extra=self.extra_dict_common)
            if msg_flg:
                return True
            else:
                return False
        except Exception as error:
            self.logger.info(f"Error: {error}", extra=self.extra_dict_common)
            return False

    # below function is not in use
    def owner_report_mail(self, htmltable=None,first_name=None):
        # if anshika says name to enter after Dear, then ask her what to do if snow returned with empty first name
        # what should be in the mail body in after Dear,
        # user {} also get first_name from SNOW API

        mnth = datetime.datetime.now().strftime('%h %Y')
        subject = f"Data Retention | Open Alerts & Questions | Nordics | {mnth}"
        mail_body = f"""
                <p>Dear Team,
                    <p>
                        Here is the list of SpeakUp cases that are due for Data redaction this month. 
                    </p>
                    <p>
                    {htmltable}
                    </p>
                    <p>
                        You need to go to the case and click on the redact button on the top right hand corner of the page and select any Personal Identifiable Information on the page that needs to be redacted and complete the action.
                        Once done you need to select ‘Yes’ in the dropdown to the question - Has personal data been redacted?
                        <b>Please note that data redaction should be done within the timelines you have mentioned so as to be compliant with your local data protection regulations.</b>                        
                    </p>
                <p style="color:#5374B8;">
                    <b>Regards,</b> <br>
                    Group Ethics Team <br>
                    Capgemini Group | Mumbai <br><br>
                    Tel: +91 22 6755 7000 (Ext. 228 6501) | Mob: + 91 98703 22330
                </p>        
                </p>"""
        return [subject, mail_body]

    # def declare_mail(self, htmltable=None, region=None):
    #
    #     mnth = datetime.datetime.now().strftime('%h %Y')
    #     subject = f"Declare | Open Declarations | {region} | {mnth}"
    #
    #     mail_body = """
    #         <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
    #           {head}
    #           <body lang="EN-US" link="#0563C1" vlink="#954F72" style="word-wrap:break-word">
    #             <div class="WordSection1">
    #               <p class="MsoNormal">
    #                 <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">Dear Team, <o:p></o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
    #                   <o:p>&nbsp;</o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">We have pulled out list of open declarations for your region; we hope this list will help you follow-up on these declarations to bring them to closure. For declarations which are open for more than 30 days (<span style="color:red">mentioned in red</span>), please make sure they are closed on priority. <o:p></o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
    #                   <o:p>&nbsp;</o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">Group Ethics provides regular updates on Declare statistics with Group Management team including Group CEO and CHRO and therefore it is important that the Declare statistics are up to date all through the year. <o:p></o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
    #                   <o:p>&nbsp;</o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
    #                   <o:p>&nbsp;</o:p>
    #                 </span>
    #               </p>
    #               {htmltable}
    #               <p class="MsoNormal">
    #                 <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
    #                   <o:p>&nbsp;</o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
    #                   <o:p>&nbsp;</o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
    #                     <b>Note:</b>  This is auto generated mail please do not reply.
    #                     Please reach out to
    #                     <a href="keith.dsouza@capgemini.com">keith.dsouza@capgemini.com</a>
    #                      for any questions/clarifications.
    #                 <o:p>&nbsp;</o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span lang="EN-GB" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#1F497D">
    #                   <o:p>&nbsp;</o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:navy">_____________________________________________________________</span>
    #                 <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif;color:navy">
    #                   <o:p></o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span lang="EN-GB" style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:#1F497D">
    #                   <o:p>&nbsp;</o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <![if !vml]>
    #                 <img width="90" height="43" style="width:.9416in;height:.45in" src="cid:sign.png" align="left" hspace="12" v:shapes="Picture_x0020_4">
    #                 <![endif]>
    #                 <b>
    #                   <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif;color:#0070AD">Group Ethics</span>
    #                 </b>
    #                 <b>
    #                   <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif;color:#0070AD">
    #                     <o:p></o:p>
    #                   </span>
    #                 </b>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span style="font-size:8.0pt;font-family:&quot;Verdana&quot;,sans-serif;color:#0070AD">Capgemini&nbsp;Group<o:p></o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span style="font-size:8.0pt;font-family:&quot;Verdana&quot;,sans-serif;color:#0070AD">
    #                   <o:p></o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <span style="font-size:10.0pt;font-family:&quot;Arial&quot;,sans-serif;color:navy">_____________________________________________________________ <o:p></o:p>
    #                 </span>
    #               </p>
    #               <p class="MsoNormal">
    #                 <o:p>&nbsp;</o:p>
    #               </p>
    #             </div>
    #           </body>
    #         </html>""".format(head=HEAD, htmltable=htmltable)
    #
    #     return [subject, mail_body]

    def speakup_mail(self, htmltable=None, region=None):

        mnth = datetime.datetime.now().strftime('%h %Y')
        subject = f"Data Retention | {region} | {mnth}"

        mail_body = f"""<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
                  {HEADER}
                  <body lang="EN-US" link="#0563C1" vlink="#954F72" style="tab-interval:.5in;word-wrap:break-word">
                    <div class="WordSection1">
                      <p class="MsoNormal">
                        <span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">Dear Team, <o:p></o:p>
                        </span>
                      </p>
                      <p class="MsoNormal">
                        <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                          <o:p>&nbsp;</o:p>
                        </span>
                      </p>
                      <p class="MsoNormal">
                        <span style="font-size:11.0pt;font-family:&quot;Calibri&quot;,sans-serif">Here is the list of SpeakUp cases that are due for Data redaction this month.<o:p></o:p>
                        </span>
                      </p>
                      <p class="MsoNormal">
                        <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                          <o:p>&nbsp;</o:p>
                        </span>
                      </p>
                      <p class="MsoNormal">
                        <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif"> <o:p></o:p>
                        </span>
                      </p>                      
                        </span>
                      </p>
                      {htmltable}
                      {CONTENT}
                    </div>
                  </body>
                </html>"""

        return [subject, mail_body]

    def declare_success(self):

        subject = "Declare | Follow-up emails | Report status | BOT process successful"
        mail_body = f"""
            <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
              {HEAD}
              <body lang="EN-US" link="#0563C1" vlink="#954F72" style="word-wrap:break-word">
                <div class="WordSection1">
                  <p class="MsoNormal">
                    <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">Hi Shantanu, <o:p></o:p>
                    </span>
                  </p>
                  <p class="MsoNormal">
                    <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                      <o:p>&nbsp;</o:p>
                    </span>
                  </p>
                  <p class="MsoNormal">
                    <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                        Declare follow-up emails have been successfully sent to the local 
                        Ethics team(s) as of {datetime.datetime.now().strftime('%d-%m-%Y')}. <o:p></o:p>
                    </span>
                  </p>
                  <p class="MsoNormal">
                    <span style="font-size:9.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                      <o:p>&nbsp;</o:p>
                    </span>
                  </p>
                  {SIGNATURE_SUPPORT}
                </div>
              </body>
            </html>     
        """
        return [subject, mail_body]

    def speakup_success(self):
        subject = "Data retention Timelines | Report status | BOT process successful"
        mail_body = f"""
                <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
                  {HEAD}
                  <body lang="EN-US" link="#0563C1" vlink="#954F72" style="word-wrap:break-word">
                    <div class="WordSection1">
                      <p class="MsoNormal">
                        <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">Hi Team, <o:p></o:p>
                        </span>
                      </p>
                      <p class="MsoNormal">
                        <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                          <o:p>&nbsp;</o:p>
                        </span>
                      </p>
                      <p class="MsoNormal">
                        <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                            Data retention Timelines emails have been successfully sent to the Issue owners and local 
                            Ethics team(s) as of  {datetime.datetime.now().strftime('%d-%m-%Y')}. <o:p></o:p>
                        </span>
                      </p>
                      <p class="MsoNormal">
                        <span style="font-size:9.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                          <o:p>&nbsp;</o:p>
                        </span>
                      </p>
                      {SIGNATURE_SUPPORT}
                    </div>
                  </body>
                </html>     
            """
        return [subject, mail_body]

    def file_missing_declare(self, errors=''):

        subject = f"Data Retention | BOT process failed"
        today = datetime.datetime.now().strftime('%d-%m-%Y')
        mail_body = f"""
            <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
              {HEAD}
              <body lang="EN-US" link="#0563C1" vlink="#954F72" style="word-wrap:break-word">
                <div class="WordSection1">
                  <p class="MsoNormal">
                    <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">Hi Team, <o:p></o:p>
                    </span>
                  </p>
                  <p class="MsoNormal">
                    <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                      <o:p>&nbsp;</o:p>
                    </span>
                  </p>
                  <p class="MsoNormal">
                    <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                        This is to inform you that the BOT process failed to execute the Declare follow-up emails 
                        to concerned Ethics team(s) as of {today}. <br><br>
                        <b><u>Error:</u></b> <br>
                        <span style="font-size:9.8pt;font-family:&quot;Verdana&quot;sans-serif">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{errors}
                        </span>
                    </span>
                  </p>
                  <p class="MsoNormal">
                    <span style="font-size:9.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                      <o:p>&nbsp;</o:p>
                    </span>
                  </p>
                  {SIGNATURE_SUPPORT}
                </div>
              </body>
            </html> 
        """
        return [subject, mail_body]

    def drt_file_missing(self, errors=''):

        subject = f"Data Retention timeline | Report Status | BOT process failed"
        today = datetime.datetime.now().strftime('%d-%m-%Y')
        mail_body = f"""
            <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
              {HEAD}
              <body lang="EN-US" link="#0563C1" vlink="#954F72" style="word-wrap:break-word">
                <div class="WordSection1">
                  <p class="MsoNormal">
                    <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">Hi Team, <o:p></o:p>
                    </span>
                  </p>
                  <p class="MsoNormal">
                    <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                      <o:p>&nbsp;</o:p>
                    </span>
                  </p>
                  <p class="MsoNormal">
                    <span style="font-size:10.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                        This is to inform you that the BOT process failed to execute the SpeakUp follow-up emails 
                        to concerned Ethics team(s) as of {today}. <br><br>
                        <b><u>Error:</u></b> <br>
                        <span style="font-size:9.8pt;font-family:&quot;Verdana&quot;sans-serif">
                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{errors}
                        </span>
                    </span>
                  </p>
                  <p class="MsoNormal">
                    <span style="font-size:9.0pt;font-family:&quot;Verdana&quot;,sans-serif">
                      <o:p>&nbsp;</o:p>
                    </span>
                  </p>
                  {SIGNATURE_SUPPORT}
                </div>
              </body>
            </html> 
        """
        return [subject, mail_body]
