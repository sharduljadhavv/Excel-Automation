from flask_app import app

@app.template_filter('datetimeformat')
def datetimeformat(value, format='%Y-%m-%d'):
    try:
        return value.strftime(format)
    except Exception as e:
        print(f"str{e}")