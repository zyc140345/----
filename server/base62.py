import base64

# base62 编码不使用在 URL 中有特殊含义的 + 和 /，因此不需要进行 URL 编码

def base62_encode(text: str) -> str:
    base64_bytes = base64.b64encode(text.encode('utf-8'))
    base62_bytes = base64_bytes.replace(b'+', b'-').replace(b'/', b'_')
    base62_bytes = base62_bytes.rstrip(b'=')
    return base62_bytes.decode('utf-8')


def base62_decode(text_base62) -> str:
    base64_str = text_base62.replace('-', '+').replace('_', '/')
    padding_needed = len(base64_str) % 4
    if padding_needed:
        base64_str += '=' * (4 - padding_needed)
    base64_bytes = base64.b64decode(base64_str)
    return base64_bytes.decode('utf-8')
