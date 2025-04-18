import win32com.client
import flask
from flask import Flask, render_template, request, session, redirect, url_for, flash, jsonify
import sqlite3
import os
import logging

# Thiết lập logging
logging.basicConfig(filename="dcom_web_client.log", level=logging.INFO, 
                   format="%(asctime)s - %(levelname)s - %(message)s")

# Khởi tạo kết nối với DCOM server
try:
    dcom_obj = win32com.client.Dispatch("DCOM.Server")
    logging.info("Kết nối thành công đến DCOM.Server")
except Exception as e:
    logging.error(f"Lỗi kết nối đến DCOM.Server: {str(e)}")
    dcom_obj = None

# Khởi tạo Flask app
app = Flask(__name__)
app.secret_key = os.urandom(24)  # Khóa bí mật cho session

# Khởi tạo kết nối đến cơ sở dữ liệu
def get_db_connection():
    conn = sqlite3.connect('customer_service.db')
    conn.row_factory = sqlite3.Row
    return conn

# Route trang chủ
@app.route('/')
def index():
    if dcom_obj is None:
        flash("Không thể kết nối đến DCOM server", "danger")
        return render_template('index.html', error=True)
    
    # Gọi phương thức hello từ DCOM server
    try:
        welcome_msg = dcom_obj.hello()
        logging.info(f"Nhận được thông điệp từ DCOM server: {welcome_msg}")
    except Exception as e:
        logging.error(f"Lỗi khi gọi phương thức hello: {str(e)}")
        welcome_msg = "Không thể lấy thông điệp chào từ DCOM server."
    
    return render_template('index.html', welcome_msg=welcome_msg)

# Route xem tất cả người dùng
@app.route('/users')
def get_all_users():
    if dcom_obj is None:
        flash("Không thể kết nối đến DCOM server", "danger")
        return redirect(url_for('index'))
    
    try:
        # Gọi phương thức fetch_all_users từ DCOM server
        users_str = dcom_obj.fetch_all_users(0)  # 0 là id mẫu
        
        # Chuyển đổi chuỗi trả về thành danh sách người dùng
        users_str = users_str.strip('[]')
        users_list = []
        
        if users_str:
            # Phân tích chuỗi từ DCOM server
            user_tuples = users_str.split('), (')
            for user_tuple in user_tuples:
                user_tuple = user_tuple.strip('()').replace("'", "")
                user_parts = user_tuple.split(', ')
                if len(user_parts) >= 2:
                    user_id, user_name = user_parts[0], user_parts[1]
                    users_list.append({'id': user_id, 'name': user_name})
        
        logging.info(f"Lấy danh sách {len(users_list)} người dùng từ DCOM server")
        return render_template('users.html', users=users_list)
    
    except Exception as e:
        logging.error(f"Lỗi khi lấy danh sách người dùng: {str(e)}")
        flash(f"Lỗi: {str(e)}", "danger")
        return redirect(url_for('index'))

# Route xem thông tin người dùng cụ thể
@app.route('/users/<int:user_id>')
def get_user(user_id):
    if dcom_obj is None:
        flash("Không thể kết nối đến DCOM server", "danger")
        return redirect(url_for('index'))
    
    try:
        # Gọi phương thức fetch_user từ DCOM server
        user_str = dcom_obj.fetch_user(user_id)
        
        if user_str == "None":
            flash(f"Không tìm thấy người dùng với ID {user_id}", "warning")
            return redirect(url_for('get_all_users'))
        
        # Chuyển đổi chuỗi trả về thành thông tin người dùng
        user_str = user_str.strip('()')
        user_parts = user_str.split(', ')
        
        if len(user_parts) >= 2:
            user_id = user_parts[0]
            user_name = user_parts[1].strip("'")
            user = {'id': user_id, 'name': user_name}
            return render_template('user_detail.html', user=user)
        else:
            flash("Định dạng dữ liệu người dùng không hợp lệ", "warning")
            return redirect(url_for('get_all_users'))
    
    except Exception as e:
        logging.error(f"Lỗi khi lấy thông tin người dùng {user_id}: {str(e)}")
        flash(f"Lỗi: {str(e)}", "danger")
        return redirect(url_for('get_all_users'))

# Route xem yêu cầu
@app.route('/requests')
def get_requests():
    if dcom_obj is None:
        flash("Không thể kết nối đến DCOM server", "danger")
        return redirect(url_for('index'))
    
    try:
        # Gọi phương thức fetch_request từ DCOM server
        request_str = dcom_obj.fetch_request(0)  # 0 là id mẫu
        
        if request_str == "None":
            flash("Không có yêu cầu nào", "info")
            return render_template('requests.html', requests=[])
        
        # Chuyển đổi chuỗi trả về thành thông tin yêu cầu
        request_str = request_str.strip('()')
        request_parts = request_str.split(', ')
        
        if len(request_parts) >= 3:
            request_id = request_parts[0]
            request_type = request_parts[1].strip("'")
            request_detail = request_parts[2].strip("'")
            request_data = {'id': request_id, 'type': request_type, 'detail': request_detail}
            return render_template('requests.html', requests=[request_data])
        else:
            flash("Định dạng dữ liệu yêu cầu không hợp lệ", "warning")
            return render_template('requests.html', requests=[])
    
    except Exception as e:
        logging.error(f"Lỗi khi lấy danh sách yêu cầu: {str(e)}")
        flash(f"Lỗi: {str(e)}", "danger")
        return render_template('requests.html', requests=[])

# API route để kiểm tra trạng thái DCOM server
@app.route('/api/status')
def api_status():
    if dcom_obj is None:
        return jsonify({"status": "error", "message": "DCOM server không khả dụng"})
    try:
        msg = dcom_obj.hello()
        return jsonify({"status": "success", "message": msg})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

# Xử lý lỗi 404
@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'), 404

# Xử lý lỗi 500
@app.errorhandler(500)
def internal_server_error(e):
    logging.error(f"Lỗi server: {str(e)}")
    return render_template('500.html'), 500

if __name__ == '__main__':
    # Tạo thư mục templates nếu chưa tồn tại
    if not os.path.exists('templates'):
        os.makedirs('templates')
    
    # Tạo các template cơ bản nếu chưa tồn tại
    template_files = ['index.html', 'users.html', 'user_detail.html', 'requests.html', '404.html', '500.html']
    for template in template_files:
        if not os.path.exists(f'templates/{template}'):
            with open(f'templates/{template}', 'w', encoding='utf-8') as f:
                f.write(f'<!-- {template} -->\n<!DOCTYPE html>\n<html>\n<head>\n  <title>DCOM Web Client</title>\n</head>\n<body>\n  <h1>DCOM Web Client</h1>\n  <p>This is a placeholder for {template}</p>\n</body>\n</html>')
    
    # Chạy ứng dụng Flask
    app.run(debug=True,port=5000)