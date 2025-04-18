from datetime import datetime


class RequestStatus(enumerate):
    NEW=1
    ANALYZING=2
    PROCESSING=3
    RESOLVED=4
    CLOSED=5
class RequestPriority(enumerate):
    LOW=1
    MEDIUM=2
    HIGH=3
    URGENT=4
class Requests:
    id:str
    timestamp:datetime
    client_id:str
    description:str
    category:str
    status:RequestStatus
    priority:RequestPriority
class RequestReceiver:
    def receive_request(self,request:Requests):

