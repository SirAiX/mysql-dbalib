# Demo Exam Super Cheatsheet: MySQL + Python + Tkinter

Используй этот файл как учебную памятку и как разрешенный справочник только если правила экзамена допускают личные материалы. Цель не в копировании готового решения, а в быстром восстановлении каркаса: БД, SQL, репозитории, Tkinter-экраны, CRUD, импорт и отладка.

## 0. Главная стратегия

Не делай идеальное приложение. Закрывай оцениваемые пункты.

Порядок, который дает больше всего баллов быстрее всего:

```text
1. БД: таблицы, PK, FK, импорт, SQL dump, ERD PDF
2. Авторизация: логин, роли, выход
3. Товары: список из БД
4. Поиск + фильтр + сортировка в реальном времени
5. CRUD товара + запрет удаления товара из заказа
6. Заказы: список, форма, CRUD
7. Документы: блок-схема PDF, скриншоты DOCX
```

Минимальная философия:

```text
Сначала работает.
Потом не падает.
Потом выглядит похоже.
Потом красиво.
```

## 1. Структура проекта

```text
project/
  app/
    main.py
    config.py
    database.py
    repositories.py
    ui.py
    assets/
      icon.ico
      icon.png
      picture.png
      products/
  database/
    schema.sql
    dump.sql
  docs/
    ERD.pdf
    flowchart.pdf
    screenshots.docx
  tools/
    smoke_test.py
  run.bat
```

## 2. MySQL: подключение

CLI с хоста:

```powershell
mysql -h 127.0.0.1 -P 3308 -u vshp -p vshp_var3
```

Через Docker:

```powershell
cd D:\Dev\Projects\vshp\my-dev-var3
docker-compose exec mysql mysql -u vshp -p vshp_var3
```

Проверка таблиц:

```sql
SHOW TABLES;
SELECT COUNT(*) FROM products;
SELECT COUNT(*) FROM orders;
SELECT COUNT(*) FROM order_items;
```

## 3. MySQL: правильные имена

Не используй зарезервированные/неудачные имена:

```text
Плохо: user, order, date, status
Хорошо: users, orders, order_date, order_statuses
```

Именование:

```text
таблицы: plural snake_case       products, order_items
поля: snake_case                 pickup_point_id
PK: id
FK: table_id                     product_id, order_id
справочники: id, name
```

## 4. MySQL: схема в 3НФ

Базовый каркас:

```sql
CREATE TABLE roles (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(64) NOT NULL UNIQUE
);

CREATE TABLE users (
    id INT AUTO_INCREMENT PRIMARY KEY,
    role_id INT NOT NULL,
    login VARCHAR(128) NOT NULL UNIQUE,
    password VARCHAR(255) NOT NULL,
    full_name VARCHAR(255) NOT NULL,
    FOREIGN KEY (role_id) REFERENCES roles(id)
);

CREATE TABLE categories (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(128) NOT NULL UNIQUE
);

CREATE TABLE suppliers (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(128) NOT NULL UNIQUE
);

CREATE TABLE manufacturers (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(128) NOT NULL UNIQUE
);

CREATE TABLE units (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(64) NOT NULL UNIQUE
);

CREATE TABLE products (
    id INT AUTO_INCREMENT PRIMARY KEY,
    sku VARCHAR(64) NOT NULL UNIQUE,
    name VARCHAR(255) NOT NULL,
    unit_id INT NOT NULL,
    price DECIMAL(10,2) NOT NULL CHECK (price >= 0),
    supplier_id INT NOT NULL,
    manufacturer_id INT NOT NULL,
    category_id INT NOT NULL,
    sale_percent DECIMAL(5,2) NOT NULL DEFAULT 0 CHECK (sale_percent >= 0 AND sale_percent <= 100),
    balance INT NOT NULL DEFAULT 0 CHECK (balance >= 0),
    description TEXT,
    photo VARCHAR(512),
    FOREIGN KEY (unit_id) REFERENCES units(id),
    FOREIGN KEY (supplier_id) REFERENCES suppliers(id),
    FOREIGN KEY (manufacturer_id) REFERENCES manufacturers(id),
    FOREIGN KEY (category_id) REFERENCES categories(id)
);

CREATE TABLE pickup_points (
    id INT PRIMARY KEY,
    address VARCHAR(255) NOT NULL UNIQUE
);

CREATE TABLE order_statuses (
    id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(128) NOT NULL UNIQUE
);

CREATE TABLE orders (
    id INT PRIMARY KEY,
    order_date DATE,
    delivery_date DATE,
    pickup_point_id INT NOT NULL,
    customer_name VARCHAR(255),
    pickup_code VARCHAR(32),
    status_id INT NOT NULL,
    FOREIGN KEY (pickup_point_id) REFERENCES pickup_points(id),
    FOREIGN KEY (status_id) REFERENCES order_statuses(id)
);

CREATE TABLE order_items (
    id INT AUTO_INCREMENT PRIMARY KEY,
    order_id INT NOT NULL,
    product_id INT NOT NULL,
    quantity INT NOT NULL CHECK (quantity > 0),
    FOREIGN KEY (order_id) REFERENCES orders(id) ON DELETE CASCADE,
    FOREIGN KEY (product_id) REFERENCES products(id) ON DELETE RESTRICT,
    UNIQUE (order_id, product_id)
);
```

3НФ в одну фразу:

```text
В основной таблице не хранить текст справочника, хранить *_id.
```

## 5. SQL: заполнить FK по текстовому значению

Сначала справочник:

```sql
INSERT IGNORE INTO categories(name)
SELECT DISTINCT category
FROM import_products
WHERE category IS NOT NULL AND category <> '';
```

Потом перенос в нормальную таблицу:

```sql
INSERT INTO products(
    sku, name, unit_id, price, supplier_id, manufacturer_id,
    category_id, sale_percent, balance, description, photo
)
SELECT
    ip.sku,
    ip.name,
    u.id,
    ip.price,
    s.id,
    m.id,
    c.id,
    ip.sale,
    ip.quantity,
    ip.description,
    ip.photo
FROM import_products ip
JOIN units u ON u.name = ip.unit
JOIN suppliers s ON s.name = ip.supplier
JOIN manufacturers m ON m.name = ip.vendor
JOIN categories c ON c.name = ip.category
ON DUPLICATE KEY UPDATE
    name = VALUES(name),
    price = VALUES(price),
    sale_percent = VALUES(sale_percent),
    balance = VALUES(balance),
    description = VALUES(description),
    photo = VALUES(photo);
```

Если уже есть текстовые поля и надо заполнить `*_id`:

```sql
UPDATE products p
JOIN categories c ON c.name = p.category
SET p.category_id = c.id;
```

## 6. CSV import через MySQL

Сначала staging-таблица:

```sql
CREATE TABLE import_products (
    sku VARCHAR(64),
    name VARCHAR(255),
    unit VARCHAR(64),
    price DECIMAL(10,2),
    supplier VARCHAR(128),
    vendor VARCHAR(128),
    category VARCHAR(128),
    sale DECIMAL(5,2),
    quantity INT,
    description TEXT,
    photo VARCHAR(512)
);
```

Загрузить CSV:

```sql
LOAD DATA INFILE '/var/lib/mysql-files/Tovar.csv'
INTO TABLE import_products
CHARACTER SET utf8mb4
FIELDS TERMINATED BY ';'
ENCLOSED BY '"'
LINES TERMINATED BY '\r\n'
IGNORE 1 ROWS
(sku, name, unit, price, supplier, vendor, category, sale, quantity, description, photo);
```

Скопировать CSV в контейнер:

```powershell
docker cp "D:\path\Tovar.csv" my-dev-var3-mysql:/var/lib/mysql-files/Tovar.csv
```

## 7. SKUS: разложить заказ

Импортная строка:

```text
PMEZMH, 2, BPV4MM, 2
```

Должна стать:

```text
order_items:
order_id | product_id | quantity
1        | id(PMEZMH) | 2
1        | id(BPV4MM) | 2
```

Python-парсер:

```python
def parse_skus(value: str) -> list[tuple[str, int]]:
    parts = [part.strip() for part in value.split(",") if part.strip()]

    items = []
    for index in range(0, len(parts), 2):
        sku = parts[index]
        quantity = int(parts[index + 1])
        items.append((sku, quantity))

    return items
```

## 8. config.py

```python
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
APP_DIR = PROJECT_ROOT / "app"
ASSETS_DIR = APP_DIR / "assets"
PRODUCT_IMAGES_DIR = ASSETS_DIR / "products"

APP_TITLE = "СтройМатериалы"
FONT = "Calibri"

COLOR_WHITE = "#FFFFFF"
COLOR_SECONDARY = "#DAA520"
COLOR_ACCENT = "#B8860B"
COLOR_BIG_SALE = "#F4A460"
COLOR_EMPTY_STOCK = "#BDEBFF"

ROLE_GUEST = "guest"
ROLE_CLIENT = "client"
ROLE_MANAGER = "manager"
ROLE_ADMIN = "admin"

DB_CONFIG = {
    "host": "127.0.0.1",
    "port": 3308,
    "user": "vshp",
    "password": "vshp",
    "database": "vshp_var3",
}
```

## 9. database.py

Если драйвер `mysql.connector` разрешен/доступен:

```python
import mysql.connector

from config import DB_CONFIG


def get_connection():
    return mysql.connector.connect(
        host=DB_CONFIG["host"],
        port=DB_CONFIG["port"],
        user=DB_CONFIG["user"],
        password=DB_CONFIG["password"],
        database=DB_CONFIG["database"],
        charset="utf8mb4",
    )
```

Важно:

```python
cursor = connection.cursor(dictionary=True)
```

Так строки можно читать как словари:

```python
row["id"]
row["name"]
```

Если строго только встроенный Python, прямого MySQL-драйвера нет. Тогда только `subprocess` и внешний `mysql`, но это медленнее и хуже для GUI.

## 10. repositories.py: авторизация

```python
from dataclasses import dataclass

from database import get_connection


@dataclass
class CurrentUser:
    id: int | None
    full_name: str
    role_name: str


class AuthRepository:
    def authenticate(self, login: str, password: str) -> CurrentUser | None:
        with get_connection() as connection:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT users.id, users.full_name, roles.name AS role_name
                FROM users
                JOIN roles ON roles.id = users.role_id
                WHERE users.login = %s AND users.password = %s
                """,
                (login, password),
            )
            row = cursor.fetchone()

        if not row:
            return None

        return CurrentUser(row["id"], row["full_name"], row["role_name"])
```

Ошибка, которую нельзя допускать:

```python
print(row["id"])  # плохо до проверки if not row
```

Правильно:

```python
if not row:
    return None
```

## 11. repositories.py: список товаров

```python
class ProductRepository:
    SORT_FIELDS = {
        "Без сортировки": "products.id",
        "Цена": "products.price",
        "Количество": "products.balance",
        "Скидка": "products.sale_percent",
    }

    def list_products(
        self,
        search: str = "",
        manufacturer_id: int | None = None,
        sort_field: str = "Без сортировки",
        sort_direction: str = "ASC",
    ) -> list[dict]:
        query = """
            SELECT
                products.id,
                products.sku,
                products.name,
                units.name AS unit_name,
                products.price,
                suppliers.name AS supplier_name,
                manufacturers.name AS manufacturer_name,
                categories.name AS category_name,
                products.sale_percent,
                products.balance,
                products.description,
                products.photo,
                ROUND(products.price * (1 - products.sale_percent / 100), 2) AS final_price
            FROM products
            JOIN units ON units.id = products.unit_id
            JOIN suppliers ON suppliers.id = products.supplier_id
            JOIN manufacturers ON manufacturers.id = products.manufacturer_id
            JOIN categories ON categories.id = products.category_id
            WHERE 1 = 1
        """
        params = []

        if search:
            like = f"%{search}%"
            query += """
                AND (
                    products.sku LIKE %s OR
                    products.name LIKE %s OR
                    products.description LIKE %s OR
                    manufacturers.name LIKE %s OR
                    suppliers.name LIKE %s OR
                    categories.name LIKE %s
                )
            """
            params.extend([like, like, like, like, like, like])

        if manufacturer_id:
            query += " AND products.manufacturer_id = %s"
            params.append(manufacturer_id)

        order_column = self.SORT_FIELDS.get(sort_field, "products.id")
        direction = "DESC" if sort_direction == "DESC" else "ASC"
        query += f" ORDER BY {order_column} {direction}"

        with get_connection() as connection:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(query, params)
            return cursor.fetchall()
```

## 12. repositories.py: CRUD товара

```python
    def get_product(self, product_id: int) -> dict | None:
        with get_connection() as connection:
            cursor = connection.cursor(dictionary=True)
            cursor.execute("SELECT * FROM products WHERE id = %s", (product_id,))
            return cursor.fetchone()

    def save_product(self, data: dict) -> int:
        self._validate_product(data)

        with get_connection() as connection:
            cursor = connection.cursor()

            if data.get("id"):
                cursor.execute(
                    """
                    UPDATE products
                    SET sku=%s, name=%s, unit_id=%s, price=%s,
                        supplier_id=%s, manufacturer_id=%s, category_id=%s,
                        sale_percent=%s, balance=%s, description=%s, photo=%s
                    WHERE id=%s
                    """,
                    (
                        data["sku"], data["name"], data["unit_id"], data["price"],
                        data["supplier_id"], data["manufacturer_id"], data["category_id"],
                        data["sale_percent"], data["balance"], data["description"],
                        data["photo"], data["id"],
                    ),
                )
                product_id = data["id"]
            else:
                cursor.execute(
                    """
                    INSERT INTO products(
                        sku, name, unit_id, price, supplier_id, manufacturer_id,
                        category_id, sale_percent, balance, description, photo
                    )
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    """,
                    (
                        data["sku"], data["name"], data["unit_id"], data["price"],
                        data["supplier_id"], data["manufacturer_id"], data["category_id"],
                        data["sale_percent"], data["balance"], data["description"], data["photo"],
                    ),
                )
                product_id = cursor.lastrowid

            connection.commit()
            return product_id

    def delete_product(self, product_id: int) -> None:
        with get_connection() as connection:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                "SELECT COUNT(*) AS count FROM order_items WHERE product_id = %s",
                (product_id,),
            )
            row = cursor.fetchone()
            if row["count"] > 0:
                raise ValueError("Товар присутствует в заказе, удалить нельзя.")

            cursor.execute("DELETE FROM products WHERE id = %s", (product_id,))
            connection.commit()

    def _validate_product(self, data: dict) -> None:
        if not data["sku"]:
            raise ValueError("Заполните артикул.")
        if not data["name"]:
            raise ValueError("Заполните наименование.")
        if data["price"] < 0:
            raise ValueError("Цена не может быть отрицательной.")
        if data["balance"] < 0:
            raise ValueError("Остаток не может быть отрицательным.")
        if data["sale_percent"] < 0 or data["sale_percent"] > 100:
            raise ValueError("Скидка должна быть от 0 до 100.")
```

## 13. ui.py: мантра Tkinter

```text
Tk создает окно.
Frame создает экран.
StringVar хранит ввод.
Entry пишет в StringVar.
Button вызывает command.
Treeview показывает список.
Toplevel открывает форму.
messagebox спасает от падения.
refresh обновляет данные.
```

Команда у кнопки без скобок:

```python
tk.Button(parent, text="Сохранить", command=self._save)
```

Плохо:

```python
command=self._save()
```

## 14. ui.py: Application

```python
import tkinter as tk
from tkinter import ttk, messagebox

import config
from repositories import AuthRepository, CurrentUser, ProductRepository


class Application:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.current_frame: tk.Frame | None = None
        self.product_form_window = None
        self.current_user: CurrentUser | None = None

        self.root.title(config.APP_TITLE)
        self.root.geometry("1200x800")
        self.root.configure(bg=config.COLOR_WHITE)
        ttk.Style().configure(".", font=(config.FONT, 10))

    def show_login(self) -> None:
        self.current_user = None
        self._show_frame(LoginFrame(self.root, self))

    def show_products(self, user: CurrentUser) -> None:
        self.current_user = user
        self._show_frame(ProductListFrame(self.root, self, user))

    def _show_frame(self, frame: tk.Frame) -> None:
        if self.current_frame:
            self.current_frame.destroy()
        self.current_frame = frame
        self.current_frame.pack(fill="both", expand=True)
```

## 15. ui.py: LoginFrame

```python
class LoginFrame(tk.Frame):
    def __init__(self, master: tk.Misc, app: Application) -> None:
        super().__init__(master, bg=config.COLOR_WHITE)
        self.app = app
        self.auth_repository = AuthRepository()
        self.login_var = tk.StringVar()
        self.password_var = tk.StringVar()
        self._build()

    def _build(self) -> None:
        form = tk.Frame(self, bg=config.COLOR_WHITE)
        form.place(relx=0.5, rely=0.5, anchor="center")

        tk.Label(form, text="Вход", bg=config.COLOR_WHITE, font=(config.FONT, 22, "bold")).grid(row=0, column=0, columnspan=2)
        tk.Label(form, text="Логин", bg=config.COLOR_WHITE).grid(row=1, column=0, sticky="w")
        tk.Entry(form, textvariable=self.login_var, width=34).grid(row=1, column=1)
        tk.Label(form, text="Пароль", bg=config.COLOR_WHITE).grid(row=2, column=0, sticky="w")
        tk.Entry(form, textvariable=self.password_var, show="*", width=34).grid(row=2, column=1)

        tk.Button(form, text="Войти", command=self._login).grid(row=3, column=0, columnspan=2, sticky="ew")
        tk.Button(form, text="Гость", command=self._guest_login).grid(row=4, column=0, columnspan=2, sticky="ew")

    def _login(self) -> None:
        try:
            user = self.auth_repository.authenticate(self.login_var.get(), self.password_var.get())
            if not user:
                messagebox.showerror("Ошибка", "Неверный логин или пароль.")
                return
            self.app.show_products(user)
        except Exception as error:
            messagebox.showerror("Ошибка", str(error))

    def _guest_login(self) -> None:
        self.app.show_products(CurrentUser(None, "Гость", config.ROLE_GUEST))
```

## 16. ui.py: ProductListFrame

```python
class ProductListFrame(tk.Frame):
    def __init__(self, master: tk.Misc, app: Application, user: CurrentUser) -> None:
        super().__init__(master, bg=config.COLOR_WHITE)
        self.app = app
        self.user = user
        self.repository = ProductRepository()

        self.search_var = tk.StringVar()
        self.manufacturer_var = tk.StringVar(value="Все производители")
        self.sort_var = tk.StringVar(value="Без сортировки")
        self.direction_var = tk.StringVar(value="ASC")

        self.can_filter = user.role_name in (config.ROLE_MANAGER, config.ROLE_ADMIN)
        self.can_admin = user.role_name == config.ROLE_ADMIN

        self._build()
        self.refresh_products()

    def _build(self) -> None:
        header = tk.Frame(self, bg=config.COLOR_SECONDARY)
        header.pack(fill="x")
        tk.Label(header, text="Товары", bg=config.COLOR_SECONDARY, fg="white", font=(config.FONT, 18, "bold")).pack(side="left", padx=10)
        tk.Label(header, text=self.user.full_name, bg=config.COLOR_SECONDARY, fg="white").pack(side="right", padx=10)
        tk.Button(header, text="Выйти", command=self.app.show_login).pack(side="right", padx=10)

        toolbar = tk.Frame(self, bg=config.COLOR_WHITE)
        toolbar.pack(fill="x", padx=10, pady=10)

        if self.can_filter:
            tk.Entry(toolbar, textvariable=self.search_var, width=30).pack(side="left", padx=5)
            ttk.Combobox(toolbar, textvariable=self.sort_var, values=["Без сортировки", "Цена", "Количество", "Скидка"], state="readonly").pack(side="left", padx=5)
            ttk.Combobox(toolbar, textvariable=self.direction_var, values=["ASC", "DESC"], state="readonly", width=6).pack(side="left", padx=5)

            for var in (self.search_var, self.sort_var, self.direction_var):
                var.trace_add("write", lambda *_: self.refresh_products())

        if self.can_admin:
            tk.Button(toolbar, text="Добавить товар", command=lambda: self.open_product_form(None)).pack(side="right", padx=5)

        columns = ("id", "sku", "name", "price", "sale_percent", "balance")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        for column in columns:
            self.tree.heading(column, text=column)
            self.tree.column(column, width=120)
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

    def refresh_products(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        products = self.repository.list_products(
            search=self.search_var.get() if self.can_filter else "",
            sort_field=self.sort_var.get() if self.can_filter else "Без сортировки",
            sort_direction=self.direction_var.get() if self.can_filter else "ASC",
        )

        for product in products:
            tag = ""
            if product["balance"] == 0:
                tag = "empty"
            elif product["sale_percent"] > 12:
                tag = "sale"

            self.tree.insert(
                "",
                "end",
                values=(
                    product["id"],
                    product["sku"],
                    product["name"],
                    product["price"],
                    product["sale_percent"],
                    product["balance"],
                ),
                tags=(tag,),
            )

        self.tree.tag_configure("sale", background=config.COLOR_BIG_SALE)
        self.tree.tag_configure("empty", background=config.COLOR_EMPTY_STOCK)
```

## 17. ui.py: Toplevel form

```python
class ProductFormWindow(tk.Toplevel):
    def __init__(self, master, product_id: int | None, refresh_callback):
        super().__init__(master)
        self.product_id = product_id
        self.refresh_callback = refresh_callback
        self.repository = ProductRepository()

        self.title("Товар")
        self.geometry("700x500")
        self.transient(master)
        self.grab_set()

        self.name_var = tk.StringVar()
        self.price_var = tk.StringVar()
        self.balance_var = tk.StringVar()
        self.sale_var = tk.StringVar(value="0")

        self._build()

        if self.product_id:
            self._load_product()

    def _build(self):
        tk.Label(self, text="Название").grid(row=0, column=0)
        tk.Entry(self, textvariable=self.name_var).grid(row=0, column=1)
        tk.Label(self, text="Цена").grid(row=1, column=0)
        tk.Entry(self, textvariable=self.price_var).grid(row=1, column=1)
        tk.Label(self, text="Остаток").grid(row=2, column=0)
        tk.Entry(self, textvariable=self.balance_var).grid(row=2, column=1)
        tk.Button(self, text="Сохранить", command=self._save).grid(row=99, column=1)

    def _save(self):
        try:
            data = {
                "id": self.product_id,
                "name": self.name_var.get().strip(),
                "price": float(self.price_var.get().replace(",", ".")),
                "balance": int(self.balance_var.get()),
                "sale_percent": float(self.sale_var.get().replace(",", ".")),
            }
            self.repository.save_product(data)
            self.refresh_callback()
            self.destroy()
        except Exception as error:
            messagebox.showerror("Ошибка", str(error))
```

## 18. OrderRepository minimum

```python
class OrderRepository:
    def list_orders(self) -> list[dict]:
        with get_connection() as connection:
            cursor = connection.cursor(dictionary=True)
            cursor.execute(
                """
                SELECT
                    orders.id,
                    order_statuses.name AS status_name,
                    pickup_points.address AS pickup_address,
                    orders.order_date,
                    orders.delivery_date,
                    orders.customer_name,
                    orders.pickup_code,
                    GROUP_CONCAT(CONCAT(products.sku, ' x', order_items.quantity) SEPARATOR ', ') AS items
                FROM orders
                JOIN pickup_points ON pickup_points.id = orders.pickup_point_id
                JOIN order_statuses ON order_statuses.id = orders.status_id
                LEFT JOIN order_items ON order_items.order_id = orders.id
                LEFT JOIN products ON products.id = order_items.product_id
                GROUP BY orders.id
                ORDER BY orders.id
                """
            )
            return cursor.fetchall()
```

## 19. Debug commands

Из папки `app`:

```powershell
py -m compileall -q .
py -c "import config, database, repositories, ui; print('IMPORT OK')"
py -c "from database import get_connection; c=get_connection(); print('DB OK'); c.close()"
py -c "from repositories import AuthRepository; print(AuthRepository().authenticate('94d5ous@gmail.com','uzWC67'))"
```

Запуск:

```powershell
py main.py
```

## 20. Python dd / var_dump

```python
from pprint import pprint


def dump(*values):
    for value in values:
        pprint(value)


def dd(*values):
    dump(*values)
    raise SystemExit
```

Использование:

```python
print(f"{row=}")
print(type(value), repr(value))
dump(products)
dd(data)
```

Встроенный debugger:

```python
breakpoint()
```

Команды:

```text
p variable
pp variable
n
s
c
q
```

## 21. Частые ошибки

### tuple indices must be integers

Причина:

```python
cursor = connection.cursor()
row["id"]
```

Фикс:

```python
cursor = connection.cursor(dictionary=True)
```

### NoneType is not subscriptable

Причина:

```python
row = cursor.fetchone()
print(row["id"])
if not row:
    return None
```

Фикс:

```python
row = cursor.fetchone()
if not row:
    return None
print(row["id"])
```

### Button command не работает

Плохо:

```python
tk.Button(root, command=self.save())
```

Хорошо:

```python
tk.Button(root, command=self.save)
```

### Tkinter падает в callback

Оборачивай методы кнопок:

```python
def _save(self):
    try:
        ...
    except Exception as error:
        messagebox.showerror("Ошибка", str(error))
```

## 22. Smoke checklist перед сдачей

```text
[ ] py -m compileall -q . проходит
[ ] приложение запускается
[ ] неверный логин показывает ошибку
[ ] админ входит
[ ] гость входит
[ ] после входа видно ФИО
[ ] товары загружаются из БД
[ ] поиск работает
[ ] сортировка работает
[ ] фильтр по производителю работает
[ ] поиск + фильтр + сортировка работают вместе
[ ] товар добавляется
[ ] товар редактируется
[ ] товар удаляется
[ ] товар из заказа не удаляется
[ ] заказы открываются у менеджера и админа
[ ] заказ добавляется
[ ] заказ редактируется
[ ] заказ удаляется
[ ] ERD сохранена в PDF
[ ] блок-схема сохранена в PDF
[ ] DOCX со скриншотами создан
[ ] SQL dump/schema.sql есть
```

## 23. Что делать, если времени мало

Осталось 60 минут:

```text
1. Убрать падения
2. Доделать список товаров
3. Доделать поиск/фильтр/сортировку
4. Добавить простую форму товара
5. Сделать ERD PDF и скриншоты
```

Осталось 30 минут:

```text
1. compileall
2. неверный логин не падает
3. товары видны из БД
4. ERD PDF
5. SQL dump
6. скриншоты того, что работает
```

Не полируй дизайн, если нет списка товаров из БД.

## 24. Главная формула экзамена

```text
БД дает фундамент.
Авторизация открывает роли.
Treeview быстро дает список.
trace_add быстро дает realtime.
Toplevel быстро дает форму.
messagebox спасает от падения.
repository держит SQL.
refresh возвращает актуальность.
```
