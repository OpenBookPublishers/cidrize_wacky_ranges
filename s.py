import httplib2
import json

def process():
    with open('data.json') as f:
        data = json.load(f)
    for institution_data in data:
        url = "http://192.168.99.100:8181/institutions"
        h = httplib2.Http()
        res, content = h.request(url,"POST",json.dumps(institution_data),
                                 {"content-type":"application/json"})
        if res.status != 200:
            print(content)

if __name__ == "__main__":
    process()
                                
