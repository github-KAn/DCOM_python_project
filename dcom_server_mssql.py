import datetime

import pythoncom
import win32com.server.util
import sqlite3
import win32api
import logging
import  request_receive_serivce

# Khởi tạo logging
logging.basicConfig(filename="dcom_server.log", level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")
# tên database
sql_db="customer_service.db"


class DCOMServer:
    _public_methods_ = ['hello','fetch_user','add_users','fetch_all_users','fetch_request']
    _reg_progid_ = "DCOM.Server"
    _reg_clsid_ = "{F9DABC68-3C3A-4F4D-8B25-65853ABEE832}"

    def hello(self):
        """ chào người dùng và lưu lại sự kiện gọi đến phương thức chào vào file dcom_server.log """
        user = win32api.GetUserName()
        message = f"Hello, {user}! Welcome to my DCOM Server."
        logging.info(f"Hello method called by {user}")
        return message
    def __init__(self):
        """ Khởi tạo kết nối database """
        self.conn = sqlite3.connect(sql_db, check_same_thread=False)
        self.cur = self.conn.cursor()
    def fetch_all_users(self, _id):
        logging.info("FetchUsers method called")
        cur=self.cur
        try:
            sql_command="SELECT id, name FROM Users "
            cur.execute(sql_command)
            users = cur.fetchall()
            return str(users)
        except Exception as e:
            logging.error(f"Database error: {str(e)}")
            return f"Error: {str(e)}"
    def fetch_user(self, _id):
        logging.info("fetch_user method called")
        cur=self.cur
        try:
            sql_command=f"SELECT id, name FROM Users where id={_id}"
            cur.execute(sql_command)
            user = cur.fetchone()
            return str(user)
        except Exception as e:
            logging.error(f"Database error: {str(e)}")
            return f"Error: {str(e)}"
    def fetch_request(self, _id):
        logging.info("fetch_requests method called")
        cur=self.cur
        try:
            sql_command=f"SELECT id,type,detail FROM Requests"
            cur.execute(sql_command)
            request = cur.fetchone()
            return str(request)
        except Exception as e:
            logging.error(f"Database error: {str(e)}")
            return f"Error: {str(e)}"
    def add_users(self):
        return 0



if __name__ == "__main__":
    from win32com.server.register import UseCommandLine
    UseCommandLine(DCOMServer)
