import yagmail
from datetime import datetime
import sys
import time
import subprocess
import logging
import schedule

# 邮件配置（请替换为你的真实信息）
FROM_EMAIL = "你的邮箱"
FROM_PASSWORD = "授权码"  # 注意：这是授权码，不是登录密码
TO_EMAIL = "接收邮箱"
SMTP_SERVER = "smtp.qq.com"
SMTP_PORT = 587

# UPS 事件检测阈值（可按需调整）
LOW_BATTERY_THRESHOLD = 20
MIN_RUNTIME_MINUTES = 5

# 定时报告配置（可按需调整）
REPORT_INTERVAL_DAYS = 2 #每几天报告一次
REPORT_TIME = "00:00"  # 每天的报告时间

# 系统配置（一般无需修改）
NUT_UPS_NAME = "SUA2200R2ICH"
CHECK_INTERVAL_SECONDS = 60
LOG_FILE = "ups_monitor.log"
UPSC_PATH = r"C:\NUT\bin\upsc.exe"

# 定义英文参数到中文描述的映射字典（完整中英文对照）
PARAMETER_MAPPING = {
    # 电池相关参数
    "battery.charge": "电池电量",
    "battery.charge.low": "低电量阈值",
    "battery.charge.warning": "电量警告阈值",
    "battery.mfr.date": "电池制造商日期",
    "battery.runtime": "电池剩余运行时间",
    "battery.runtime.low": "低运行时间阈值",
    "battery.temperature": "电池温度",
    "battery.type": "电池类型",
    "battery.voltage": "电池电压",
    "battery.voltage.nominal": "电池额定电压",

    # 设备信息
    "device.mfr": "设备制造商",
    "device.model": "设备型号",
    "device.serial": "设备序列号",
    "device.type": "设备类型",

    # 驱动信息
    "driver.name": "驱动名称",
    "driver.parameter.pollfreq": "轮询频率",
    "driver.parameter.pollinterval": "轮询间隔",
    "driver.parameter.port": "连接端口",
    "driver.version": "驱动版本",
    "driver.version.data": "驱动数据版本",
    "driver.version.internal": "内部驱动版本",

    # 输入输出参数
    "input.sensitivity": "输入敏感度",
    "input.transfer.high": "高输入转换电压",
    "input.transfer.low": "低输入转换电压",
    "input.transfer.reason": "输入转换原因",
    "input.voltage": "输入电压",
    "output.current": "输出电流",
    "output.frequency": "输出频率",
    "output.voltage": "输出电压",
    "output.voltage.nominal": "输出额定电压",

    # UPS 状态与控制
    "ups.beeper.status": "蜂鸣器状态",
    "ups.delay.shutdown": "关机延迟时间",
    "ups.delay.start": "启动延迟时间",
    "ups.firmware": "固件版本",
    "ups.firmware.aux": "辅助固件版本",
    "ups.load": "负载百分比",
    "ups.mfr": "UPS 制造商",
    "ups.mfr.date": "UPS 制造日期",
    "ups.model": "UPS 型号",
    "ups.productid": "产品 ID",
    "ups.serial": "UPS 序列号",
    "ups.status": "UPS 状态",
    "ups.test.result": "自检结果",
    "ups.timer.reboot": "重启计时器",
    "ups.timer.shutdown": "关机计时器",
    "ups.timer.start": "启动计时器",
    "ups.vendorid": "供应商 ID",
}


def setup_logging():
    """初始化日志配置"""
    logging.basicConfig(
        filename=LOG_FILE,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )


def get_ups_status():
    """获取UPS实时状态（通过NUT的upsc命令）"""
    try:
        result = subprocess.run(
            [UPSC_PATH, NUT_UPS_NAME],
            capture_output=True,
            text=True,
            timeout=10
        )
        return result.stdout
    except Exception as e:
        logging.error(f"获取UPS状态失败：{str(e)}")
        return ""


def translate_status(status):
    """将英文参数转换为中英文对照格式"""
    translated_status = ""
    for line in status.split("\n"):
        if ":" in line:
            param, value = line.split(": ", 1)
            chinese_desc = PARAMETER_MAPPING.get(param, param)  # 获取中文描述，未定义的保留英文
            translated_status += f"{param}: {chinese_desc} - {value.strip()}\n"  # 中英文对照显示
        else:
            translated_status += line + "\n"
    return translated_status


def send_email(subject, body):
    """发送邮件（封装yagmail逻辑）"""
    try:
        with yagmail.SMTP(
                user=FROM_EMAIL,
                password=FROM_PASSWORD,
                host=SMTP_SERVER,
                port=SMTP_PORT,
                smtp_starttls=True,
                smtp_ssl=False
        ) as yag:
            yag.send(to=TO_EMAIL, subject=subject, contents=body)
            logging.info(f"邮件发送成功：{subject}")
    except Exception as e:
        logging.error(f"邮件发送失败：{str(e)}")


def handle_status_change(current_status, last_status):
    """处理供电状态切换事件"""
    current_status = translate_status(current_status)
    last_status = translate_status(last_status)

    # 提取状态中的关键标识（假设状态中包含 "UPS 状态: OL" 或 "UPS 状态: OB"）
    current_ups_status = next((line for line in current_status.split("\n") if "UPS 状态:" in line), "")
    last_ups_status = next((line for line in last_status.split("\n") if "UPS 状态:" in line), "")

    if "OL" in last_ups_status and "OB" in current_ups_status:
        event_type = "市电中断，切换到电池供电"
    elif "OB" in last_ups_status and "OL" in current_ups_status:
        event_type = "市电恢复，切换回市电供电"
    else:
        return

    battery_charge = parse_battery_charge(current_status)
    runtime_minutes = parse_runtime(current_status)
    body = f"""
    【紧急】UPS供电状态切换！
    - 事件：{event_type}
    - 当前电量：{battery_charge}%
    - 剩余时间：{runtime_minutes}分钟
    - 时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    - 完整状态：
{current_status}
    """.strip()
    send_email(f"【UPS事件】{event_type}", body)


def check_low_battery_or_runtime(current_status):
    """检查低电量或剩余时间不足"""
    current_status = translate_status(current_status)
    battery_charge = parse_battery_charge(current_status)
    runtime_minutes = parse_runtime(current_status)

    if battery_charge < LOW_BATTERY_THRESHOLD or runtime_minutes < MIN_RUNTIME_MINUTES:
        body = f"""
        【警告】UPS电量/时间不足！
        - 电池电量：{battery_charge}%（阈值：{LOW_BATTERY_THRESHOLD}%）
        - 剩余时间：{runtime_minutes}分钟（阈值：{MIN_RUNTIME_MINUTES}分钟）
        - 时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
        - 完整状态：
{current_status}
        """.strip()
        send_email(f"【UPS警报】电量/时间不足", body)


def parse_battery_charge(status):
    """解析电池电量（从翻译后的状态中提取）"""
    line = next((l for l in status.split("\n") if "电池电量" in l), None)
    return int(line.split("- ")[1].strip("%")) if line else 0


def parse_runtime(status):
    """解析剩余运行时间（从翻译后的状态中提取）"""
    line = next((l for l in status.split("\n") if "电池剩余运行时间" in l), None)
    return int(line.split("- ")[1]) // 60 if line else 0


def send_full_status_report():
    """发送全量状态报告"""
    status = get_ups_status()
    status = translate_status(status)
    subject = f"【UPS全量报告】{datetime.now().strftime('%Y-%m-%d')}"
    body = f"""
    UPS 全量运行报告（自动生成）
    - 报告时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    - 设备名称：{NUT_UPS_NAME}
    - 完整状态数据（英文: 中文 - 值）：
{status}
    """.strip()
    send_email(subject, body)


def send_startup_report():
    """发送开机启动状态报告邮件"""
    status = get_ups_status()
    status = translate_status(status)
    subject = f"【UPS开机启动报告】{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    body = f"""
    UPS 开机启动状态报告：
    - 启动时间：{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    - 设备名称：{NUT_UPS_NAME}
    - 当前状态数据（英文: 中文 - 值）：
{status}
    """.strip()
    send_email(subject, body)


def simulate_low_battery():
    """模拟电池电量和运行时间低于阈值"""
    mock_status = get_ups_status()
    # 设置模拟值（电量15%，剩余时间3分钟）
    mock_status = mock_status.replace("battery.charge: 100", "battery.charge: 15")
    mock_status = mock_status.replace("battery.runtime: 6000", "battery.runtime: 180")  # 3分钟=180秒
    check_low_battery_or_runtime(mock_status)
    logging.info("模拟低电量事件已触发")


def main():
    """主函数，程序入口"""
    setup_logging()
    logging.info("UPS监控脚本启动")

    send_startup_report()
    last_status = get_ups_status()

    schedule.every(REPORT_INTERVAL_DAYS).days.at(REPORT_TIME).do(send_full_status_report)

    while True:
        current_status = get_ups_status()
        if not current_status:
            time.sleep(CHECK_INTERVAL_SECONDS)
            continue

        handle_status_change(current_status, last_status)
        check_low_battery_or_runtime(current_status)
        last_status = current_status
        schedule.run_pending()
        time.sleep(CHECK_INTERVAL_SECONDS)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[1] == "test":
            status = get_ups_status()
            status = translate_status(status)
            send_email("【UPS测试邮件】", f"测试报告：\n{status}")
        elif sys.argv[1] == "simulate_low_battery":
            simulate_low_battery()
        else:
            send_email(f"【NUT事件】{sys.argv[1]}", f"接收到未知事件：{sys.argv[1]}")
    else:
        main()
