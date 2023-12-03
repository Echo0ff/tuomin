import requests

def recognize_entities_with_hanlp_api(text, api_token):
    try:
        url = "http://comdo.hanlp.com/hanlp/v21/ner/ner"
        headers = {
            "token": api_token  # 直接使用 token 而不是 Bearer 方案
        }
        data = {
            "text": text
        }

        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()

        result = response.json()
        return result
    except Exception as e:
        return {"error": str(e)}


# 您的API token
api_token = "2f41e8c978a0452291b8a12e10ecee101701132175415token"

# 要进行实体命名识别的文本
input_text = "这是一段需要进行实体命名识别的文本，例如识别人名、地名等实体。孙中山是开国元勋吗？"

# 调用API进行实体命名识别
api_result = recognize_entities_with_hanlp_api(input_text, api_token)

# 打印识别结果
print("实体命名识别结果：")
print(api_result)
