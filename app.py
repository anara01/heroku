from flask import Flask, render_template, request
import openpyxl
app = Flask(__name__)
book = openpyxl.open('goods.xlsx', read_only=True)
sheet = book.active
goods = []
i = 1
while sheet[f'A{i}'].value:
    goods.append(sheet[f'A{i}'].value)
    i += 1
book.close()

@app.route('/')
def homepage():
    book = openpyxl.open('goods.xlsx', read_only=True)
    
    return render_template('index.html', goods=goods)
@app.route('/add/', methods=["POST"])
def add():
    good = request.form["good"]
    goods.append(good)
    book = openpyxl.open('goods.xlsx')
    sheet = book.active
    x = len(goods)
    sheet[x][0].value = good
    book.save('goods.xlsx')
    book.close()
    return """
        <h1>Инвентарь пополнен</h1>
        <a href='/'>Домой</a>
    """
if __name__ == '__main__':
    app.run(debug=True)
