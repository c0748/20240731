from datetime import datetime

current_date = datetime.now()

print(current_date)

formatted_date = current_date.strftime("%Y年%m月%d日")
print(formatted_date)
