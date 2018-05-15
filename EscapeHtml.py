import html
tables = open('./out/test.html', 'rb').read().decode('utf-16')
template = open('./template.aspx', 'rb').read().decode('utf-8')
home = template.replace("<INSERTHERE>", html.escape(tables))

with open('./out/Home.aspx', 'wb') as out:
    out.write(home.encode('utf-8'))