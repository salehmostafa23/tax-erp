import requests
import base64
import json
import os

GITHUB_REPO = 'salehmostafa23/tax-erp'
GITHUB_API = 'https://api.github.com'

def _get_token():
    try:
        import streamlit as st
        return st.secrets.get('GITHUB_TOKEN','')
    except:
        return os.environ.get('GITHUB_TOKEN','')

def _headers():
    t=_get_token()
    if not t: return None
    return {'Authorization':f'token {t}','Accept':'application/vnd.github.v3+json'}

def gh_read(path):
    h=_headers()
    if not h: return None
    try:
        r=requests.get(f'{GITHUB_API}/repos/{GITHUB_REPO}/contents/{path}',headers=h,timeout=15)
        if r.status_code==200:
            return json.loads(base64.b64decode(r.json()['content']).decode('utf-8'))
    except: pass
    return None

def gh_write(path,data):
    h=_headers()
    if not h: return False
    url=f'{GITHUB_API}/repos/{GITHUB_REPO}/contents/{path}'
    sha=None
    try:
        r=requests.get(url,headers=h,timeout=15)
        if r.status_code==200: sha=r.json()['sha']
    except: pass
    content=base64.b64encode(json.dumps(data,ensure_ascii=False,indent=2,default=str).encode('utf-8')).decode('utf-8')
    payload={'message':f'Update {path}','content':content}
    if sha: payload['sha']=sha
    try:
        r=requests.put(url,headers=h,json=payload,timeout=15)
        return r.status_code in [200,201]
    except: return False
