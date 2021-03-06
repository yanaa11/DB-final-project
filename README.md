# DB-final-project
Very simple app to manage furniture orders using standard order form. 

___

"Основа" приложения - Excel файлы "База_Меб_Компаний.xls", "Сводная.xls" (прикреплены скриншоты, чтобы проиллюстрировать идею, но не выкладывать информацию компании) и бланк заказа стандартного образца "Заказ_Орг1.xls" (как пример). При оформлении заказа в этот бланк вбивается вручную состав заказа и заказчик. При оплате заказа клиентом заказу присваивается номер, бланк заказа отправляется на фабрику, а информация из бланка вручную переписывалась в "сводную". И далее там же ослеживался "жизненный цикл" заказа. Клиенты были слабо систематизированы в еще одной excel таблице "База...". Но с заказами их необходимо было связывать вручную, что делалось не всегда. 

![Image base](https://github.com/yanaa11/DB-final-project/blob/master/fasady_app/screen/companies_base.png)
![Image svodnaya](https://github.com/yanaa11/DB-final-project/blob/master/fasady_app/screen/svodnaya.png)
![Image form](https://github.com/yanaa11/DB-final-project/blob/master/fasady_app/screen/order_form.png)

Поэтому мое приложение экспортирует информацию из бланка оплаченного заказа в базу (заменяющую собой "сводную" и "базу клиентов"), ругается если нет указанного клиента (предлагая его добавить), выдает по запросу информацио о заказе и позволяет менять его статус и добавлять оплату при необходимости.  

Графический интерфейс уродлив, но он есть. Само приложение в файле "app.py". + Есть два sql файла "create_tables" и "functions", которые создают таблицы и функции соответственно. 

В python коде для получения данных к базе делаются запросы средствами 'psycopg2' (в основном обращения к сохраненным функциям: select get_somthing(id)). 

Диаграмма в ER_Diagram.pdf
Прикрепила несколько картинок самого приложения, чтобы было понятно, о чем идет речь. 

![](https://github.com/yanaa11/DB-final-project/blob/master/fasady_app/screen/upload.png)
![](https://github.com/yanaa11/DB-final-project/blob/master/fasady_app/screen/info.png)
![](https://github.com/yanaa11/DB-final-project/blob/master/fasady_app/screen/status.png)
![](https://github.com/yanaa11/DB-final-project/blob/master/fasady_app/screen/client.png)
