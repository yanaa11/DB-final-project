

create table clients (

	name varchar(100) primary key, 
	category varchar(20) check (category in ('Частный', 'Дизайнер', 'Компания')),
	tag varchar(20) check (tag in ('Постоянный', 'Лояльный', 'Новый', 'Отказной', 'Привлекаем')), 
	city varchar(50), 
	address varchar(100),
	tel varchar(20),
	mail varchar(50),
	site varchar(100), 
	inst varchar(50),
	vk varchar(100),
	last_contact date
);

create table persons (

	name varchar(100) not null, 
	client varchar(100) references clients (name) on delete cascade,
	tel varchar(20) not null, 
	mail varchar(50),
	info varchar(200), 
	last_contact date,
	primary key (name, client)

);

create table orders (

	id int primary key,
	client varchar(100),
	person varchar(100),
	state varchar(20) check (state in ('На просчете', 'Оплачен', 'Заказан', 'В пути', 'На складе', 'Выдан')),
	pay_date date not null,
	request_date date,
	recieve_date date,
	give_date date,
	foreign key (client, person) references persons (client, name) on delete set null
);

create table ord_money (

	id int references orders (id) on delete cascade primary key,
	price numeric not null,
	prepay numeric not null,
	debt numeric not null,
	type varchar(20) check (type in ('Счет', 'Приходный ордер')), 
	doc varchar(100) not null, 
	doc_date date not null,
	bill numeric,
	delivery numeric not null,
	extra numeric not null default 0.0,
	profit numeric
);

create table ord_content (

	id int references orders (id) on delete cascade primary key,
	color varchar(100) not null,
	type varchar(10) check (type in ('Краска', 'Пленка')),
	shape varchar(50) not null,
	edge varchar(50) not null,
	msq numeric not null,
	items int not null

);

