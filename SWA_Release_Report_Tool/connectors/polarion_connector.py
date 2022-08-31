from requests import Session
from zeep import Client, Settings, Transport
from zeep.exceptions import Fault
from zeep.plugins import HistoryPlugin
import urllib3


class Polarion(object):

    def __init__(self, url):
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        self.history = HistoryPlugin()
        self.session_id = None
        self.__create_client(url)

    def connect(self, user_name="", password=""):
        self.disconnect()
        try:
            self.session_webservice.service.logIn(user_name, password)
        except Fault:
            return False
        self.__init_session_id()
        self.__update_soap_headers()
        return True

    def disconnect(self):
        self.session_webservice.service.endSession()
        self.session_id = None

    def is_connected(self):
        return self.session_webservice.service.hasSubject()

    def __create_web_service(self, url, service_name):
        _WSDL_PATH = url + "/ws/services/$$?wsdl"
        setting = Settings(strict=False)
        tr = Transport(timeout=5 * 60 * 1000, operation_timeout=5 * 60 * 1000)
        session = Session()
        session.verify = False
        tr.session = session
        service = Client(wsdl=_WSDL_PATH.replace("$$", service_name),
                         plugins=[self.history], settings=setting, transport=tr)
        return service

    def __create_client(self, url):
        self.session_webservice = self.__create_web_service(url, 'SessionWebService')
        self.security_webservice = self.__create_web_service(url, "SecurityWebService")
        self.builder_webservice = self.__create_web_service(url, 'BuilderWebService')
        self.project_webservice = self.__create_web_service(url, 'ProjectWebService')
        self.planning_webservice = self.__create_web_service(url, 'PlanningWebService')
        self.test_management_webservice = self.__create_web_service(url, 'TestManagementWebService')
        self.tracker_webservice = self.__create_web_service(url, 'TrackerWebService')

    def __init_session_id(self):
        find_url = ".//{http://ws.polarion.com/session}sessionID"
        self.session_id = self.history.last_received['envelope'].getroottree().find(find_url)

    def __update_soap_headers(self):
        if self.session_id is None:
            raise ValueError('Can\'t Find Session Id')
        self.session_webservice.set_default_soapheaders([self.session_id])
        self.security_webservice.set_default_soapheaders([self.session_id])
        self.project_webservice.set_default_soapheaders([self.session_id])
        self.planning_webservice.set_default_soapheaders([self.session_id])
        self.builder_webservice.set_default_soapheaders([self.session_id])
        self.test_management_webservice.set_default_soapheaders([self.session_id])
        self.tracker_webservice.set_default_soapheaders([self.session_id])
