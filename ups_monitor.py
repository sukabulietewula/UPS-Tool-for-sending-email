import subprocess
import time
import yagmail
from datetime import datetime
import sys


# 邮件配置信息
email_config = {
    'from_email': '你的邮箱',
    'from_password': '你的邮箱授权码',
    'to_email': '接收邮箱',
    'smtp_server': 'smtp.qq.com',#对应的SMTP服务器地址
    'smtp_port': 587
}


def get_ups_status():
    try:
        # 替换为 upsc.exe 的完整路径
        result = subprocess.run(['C:\\NUT\\bin\\upsc.exe', 'SUA2200R2ICH'], capture_output=True, text=True, check=True)
        output = result.stdout
        for line in output.splitlines():
            if 'ups.status' in line:
                return line.split(':')[1].strip()
        return None
    except subprocess.CalledProcessError as e:
        print(f"Error getting UPS status: {e}")
        return None


def get_ups_parameters():
    try:
        # 替换为 upsc.exe 的完整路径
        result = subprocess.run(['C:\\NUT\\bin\\upsc.exe', 'NUT里面你的UPS的名字'], capture_output=True, text=True, check=True)
        return result.stdout
    except subprocess.CalledProcessError as e:
        print(f"Error getting UPS parameters: {e}")
        return ""


def send_email(subject, body):
    try:
        yag = yagmail.SMTP(
            user=email_config['from_email'],
            password=email_config['from_password'],
            host=email_config['smtp_server'],
            port=email_config['smtp_port'],
            smtp_starttls=True,
            smtp_ssl=False
        )
        yag.send(to=email_config['to_email'], subject=subject, contents=body)
        print(f"[{datetime.now()}] 邮件发送成功：{subject}")
    except Exception as e:
        print(f"[{datetime.now()}] 邮件发送失败：{str(e)}")


def test_status_switch():
    current_status = get_ups_status()
    mock_status = "ONBATT" if current_status != "ONBATT" else "OL"
    subject = f"【UPS测试通知】模拟供电状态切换至 {mock_status}"
    body = f"""
    【UPS测试通知】
    - 事件类型：模拟供电状态切换
    - 发生时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    - 模拟新状态：{mock_status}
    - 旧状态：{current_status}
    - UPS 当前参数：
    {get_ups_parameters()}
    """.strip()
    send_email(subject, body)


def main():
    if len(sys.argv) > 1 and sys.argv[1] == 'test':
        test_status_switch()
        return

    last_status = get_ups_status()
    while True:
        current_status = get_ups_status()
        if current_status and current_status != last_status:
            subject = f"【UPS紧急事件】UPS 供电状态切换至 {current_status}"
            body = f"""
            【UPS状态变更通知】
            - 事件类型：UPS 供电状态切换
            - 发生时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
            - 新状态：{current_status}
            - 旧状态：{last_status}
            - UPS 当前参数：
            {get_ups_parameters()}
            """.strip()
            send_email(subject, body)
            last_status = current_status
        time.sleep(10)


if __name__ == "__main__":
    main()