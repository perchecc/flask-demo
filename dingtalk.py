import os, hmac, time, base64, hashlib, json, requests
from typing import List, Dict

WEBHOOK = os.getenv("DING_WEBHOOK", "")  # 必填
SECRET = os.getenv("DING_SECRET", "")  # 如果启用了加签，则必填
# 手机号映射：支持两种写法
# 1) 全量映射：{"员工姓名":"13800000000", ...}
# 2) 只映射主管：{"冯诚":"13800000000", ...}  员工无映射时自动跳过@，只发文本
PHONE_MAP = json.loads(os.getenv("DING_PHONE_MAP", "{}"))


def _sign():
    # 钉钉加签算法
    timestamp = str(round(time.time() * 1000))
    secret_enc = SECRET.encode("utf-8")
    string_to_sign = "{}\n{}".format(timestamp, SECRET)
    string_to_sign_enc = string_to_sign.encode("utf-8")
    hmac_code = hmac.new(
        secret_enc, string_to_sign_enc, digestmod=hashlib.sha256
    ).digest()
    sign = base64.b64encode(hmac_code).decode("utf-8")
    return timestamp, sign


def send_dingtalk_warning(rows: List[Dict]):
    """
    rows: [{'name':'胡旭冉','hours':6.5,'leader':'冯诚'}, ...]
    """
    if not WEBHOOK:
        return

    # 构造@列表
    at_phones, at_names = [], []
    content_lines = ["🔔 以下同学昨日工时不足 8h，请及时关注：", ""]
    for r in rows:
        name = r["name"]
        leader = r.get("leader", "")
        phone = PHONE_MAP.get(leader) or PHONE_MAP.get(name)
        if phone:
            at_phones.append(phone)
            at_names.append(leader or name)
        content_lines.append(f"{name} ({r['hours']}h)  主管：{leader or '无'}")

    text = "\n".join(content_lines)
    data = {
        "msgtype": "text",
        "text": {"content": text},
        "at": {"atMobiles": at_phones, "isAtAll": False},
    }

    url = WEBHOOK
    if SECRET:
        ts, sign = _sign()
        url += f"&timestamp={ts}&sign={sign}"

    resp = requests.post(url, json=data, timeout=5)
    resp.raise_for_status()
