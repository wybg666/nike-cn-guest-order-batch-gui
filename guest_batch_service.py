import json
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple


DEFERRED_FORM_FIELDS = (
    "deferredFormStatus",
    "deferredFormJobId",
    "deferredFormValidUntil",
    "deferredFormAction",
    "deferredQrCodeURL",
    "deferredFormError",
)


@dataclass
class ProxyItem:
    raw: str
    host: str
    port: str
    username: Optional[str] = None
    password: Optional[str] = None

    def as_requests_proxy(self) -> Dict[str, str]:
        if self.username is not None and self.password is not None:
            base = f"http://{self.username}:{self.password}@{self.host}:{self.port}"
        else:
            base = f"http://{self.host}:{self.port}"
        return {"http": base, "https": base}


def parse_order_phone_lines(text: str) -> List[Tuple[str, str]]:
    pairs: List[Tuple[str, str]] = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        parts = [x for x in re.split(r"[\t, ]+", line) if x]
        if len(parts) < 2:
            continue
        order_id = parts[0].strip()
        phone = parts[1].strip()
        pairs.append((order_id, phone))
    return pairs


def parse_proxies(text: str) -> List[ProxyItem]:
    items: List[ProxyItem] = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        parts = line.split(":")
        if len(parts) == 2:
            host, port = parts
            items.append(ProxyItem(raw=line, host=host, port=port))
        elif len(parts) == 4:
            host, port, user, pwd = parts
            items.append(ProxyItem(raw=line, host=host, port=port, username=user, password=pwd))
    return items


def normalize_order_row(j: dict, query_phone: str) -> dict:
    tx = j.get("transaction") or {}
    st = j.get("fullOrderStatus") or {}
    shipping_raw = j.get("shippingAddress")
    if isinstance(shipping_raw, list):
        shipping = [str(x or "").strip() for x in shipping_raw]
    elif isinstance(shipping_raw, dict):
        shipping = [
            str(
                shipping_raw.get("name")
                or shipping_raw.get("receiverName")
                or shipping_raw.get("fullName")
                or ""
            ).strip(),
            str(shipping_raw.get("address1") or shipping_raw.get("line1") or "").strip(),
            str(shipping_raw.get("address2") or shipping_raw.get("line2") or "").strip(),
            str(
                shipping_raw.get("address3")
                or shipping_raw.get("line3")
                or shipping_raw.get("district")
                or shipping_raw.get("county")
                or ""
            ).strip(),
            str(
                shipping_raw.get("region")
                or shipping_raw.get("city")
                or shipping_raw.get("state")
                or shipping_raw.get("province")
                or ""
            ).strip(),
            str(
                shipping_raw.get("phone")
                or shipping_raw.get("phoneNumber")
                or shipping_raw.get("mobile")
                or ""
            ).strip(),
        ]
    else:
        shipping = []
    shipping = (shipping + ["", "", "", "", "", ""])[:6]

    full_address = " ".join([x for x in shipping if x]).strip()
    payment_approval_id = str(
        j.get("paymentApprovalId")
        or (tx.get("paymentApprovalId") if isinstance(tx, dict) else "")
        or ((j.get("paymentInfo") or [{}])[0].get("approvalId") if isinstance(j.get("paymentInfo"), list) and (j.get("paymentInfo") or []) else "")
        or ""
    ).strip()
    payment_hold_flag = str(
        j.get("paymentHoldFlag")
        or (tx.get("paymentHoldFlag") if isinstance(tx, dict) else "")
        or ""
    ).strip()

    return {
        "orderId": j.get("orderId", ""),
        "queryPhone": query_phone,
        "paymentApprovalId": payment_approval_id,
        "paymentHoldFlag": payment_hold_flag,
        "status": st.get("status", ""),
        "subStatus": st.get("subStatus", ""),
        "orderDate": tx.get("orderDate", ""),
        "currency": j.get("currency", ""),
        "orderTotal": tx.get("orderTotal", ""),
        "chargedPrice": tx.get("chargedPrice", ""),
        "paymentStatus": j.get("nonTranslatedPaymentStatus", ""),
        "receiverName": shipping[0],
        "receiverPhone": shipping[5],
        "receiverRegion": shipping[4],
        "receiverAddressLine": shipping[1],
        "receiverAddressLine2": shipping[2],
        "receiverAddressLine3": shipping[3],
        "fullShippingAddress": full_address,
        "shippingAddressRaw": json.dumps(shipping_raw, ensure_ascii=False)
        if shipping_raw is not None
        else "",
        "deferredFormStatus": "",
        "deferredFormJobId": "",
        "deferredFormValidUntil": "",
        "deferredFormAction": "",
        "deferredQrCodeURL": "",
        "deferredFormError": "",
        "rawJson": json.dumps(j, ensure_ascii=False),
    }


def normalize_items(j: dict, query_phone: str) -> List[dict]:
    rows: List[dict] = []
    order_id = j.get("orderId", "")
    groups = j.get("group") or []
    if isinstance(groups, dict):
        groups = [groups]
    for g in groups:
        if not isinstance(g, dict):
            continue
        heading = g.get("heading", "")
        subheading = g.get("subheading", "")
        encoded_message = g.get("encodedMessage", "")
        group_actions = g.get("actions")
        group_actions_raw = json.dumps(group_actions, ensure_ascii=False) if group_actions is not None else ""
        for it in (g.get("orderItems") or []):
            if not isinstance(it, dict):
                continue
            p = it.get("product") or {}
            ls = it.get("lineItemStatus") or {}
            lt = it.get("lineItemTransaction") or {}
            estimated = it.get("nonTranslatedEstimatedDelivery") or {}
            if not isinstance(estimated, dict):
                estimated = {}
            item_actions = it.get("actions")
            rows.append(
                {
                    "orderId": order_id,
                    "queryPhone": query_phone,
                    "groupHeading": heading,
                    "groupSubheading": subheading,
                    "groupEncodedMessage": encoded_message,
                    "productTitle": p.get("title", ""),
                    "productId": p.get("productId", ""),
                    "productPid": p.get("pid", ""),
                    "productCategory": p.get("category", ""),
                    "productImage": p.get("productImage", ""),
                    "styleNumber": p.get("styleNumber", ""),
                    "color": p.get("color", ""),
                    "styleColor": p.get("styleColor", ""),
                    "gtin": p.get("gtin", ""),
                    "size": p.get("size", ""),
                    "lineStatus": ls.get("status", ""),
                    "lineSubStatus": ls.get("subStatus", ""),
                    "lineStatusColor": ls.get("statusColor", ""),
                    "lineSubStatusColor": ls.get("subStatusColor", ""),
                    "quantity": lt.get("quantity", ""),
                    "returnableQuantity": lt.get("returnableQuantity", ""),
                    "lineItemChargedPrice": lt.get("lineItemChargedPrice", ""),
                    "lineItemFees": lt.get("lineItemFees", ""),
                    "lineItemOrderDate": lt.get("orderDate", ""),
                    "estimatedDelivery": it.get("eddTranslated", ""),
                    "estimatedDeliveryMaxDateTime": estimated.get("maxDateTime", ""),
                    "estimatedDeliveryPrecision": estimated.get("precision", ""),
                    "timezoneId": it.get("nonTranslatedTimezoneId", ""),
                    "orderLineType": it.get("nonTranslatedOrderlineType", ""),
                    "linePaymentStatus": it.get("nonTranslatedPaymentStatus", ""),
                    "sizeDescription": it.get("nonTranslatedSizeDescription", ""),
                    "standardCarrierAlphaCode": it.get("standardCarrierAlphaCode", ""),
                    "returnableFlag": it.get("returnableFlag", ""),
                    "groupActionsRaw": group_actions_raw,
                    "itemActionsRaw": json.dumps(item_actions, ensure_ascii=False)
                    if item_actions is not None
                    else "",
                }
            )
    return rows


def ensure_deferred_form_fields(row: dict) -> dict:
    for key in DEFERRED_FORM_FIELDS:
        row.setdefault(key, "")
    return row


def resolve_payment_approval_id(row: dict, j: dict) -> str:
    return str(
        row.get("paymentApprovalId")
        or j.get("paymentApprovalId")
        or ((j.get("transaction") or {}).get("paymentApprovalId") if isinstance(j.get("transaction"), dict) else "")
        or ""
    ).strip()


def build_failed_row(order_id: str, query_phone: str, error: str, proxy: str, attempt: int) -> dict:
    return {
        "orderId": order_id,
        "queryPhone": query_phone,
        "error": error,
        "proxy": proxy,
        "attempt": attempt,
    }
