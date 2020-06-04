/* Simple functions used in 'app.py' to get, set and check data from DB. */


create or replace function get_last_id() returns int as
	$$
	declare
		last_id int;
	begin
		select max(id) from orders into last_id;
		if last_id is null then
			return -1;
		end if;
		return last_id;
	end;
	$$ language plpgsql;

create or replace function check_id(c_id int) returns varchar as
	$$
	declare
		c int;
	begin
		select id from orders where id = c_id into c;
		if c is null then
			return 'no';
		end if;
		return 'ok';
	end;
	$$ language plpgsql;

create or replace function check_client(client varchar) returns varchar as
	$$
	declare
		res varchar;
	begin
		select name from clients where name = client into res;
		if res is null then
			return 'no';
		end if;
		return 'yes';
	end;
	$$ language plpgsql;

create or replace function check_person(client_p varchar, person varchar) returns varchar as
	$$
	declare
		res varchar;
	begin
		select name from persons where name = person and client = client_p into res;
		if res is null then
			return 'no';
		end if;
		return 'yes';
	end;
	$$ language plpgsql;

create or replace function add_client(c_name varchar, c_category varchar, c_tag varchar, c_city varchar, c_address varchar,
c_tel varchar, c_mail varchar, c_site varchar, c_inst varchar, c_vk varchar, c_last_contact date) returns varchar as
	$$
	declare
		c varchar;
	begin
		select check_client(c_name) into c;
		if c = 'yes' then
			return 'no';
		end if;
		insert into clients (name, category, tag, city, address, tel, mail, site, inst, vk, last_contact) values
			(c_name, c_category, c_tag, c_city, c_address, c_tel, c_mail, c_site, c_inst, c_vk, c_last_contact);
		return 'ok';
		exception
			when others then
    			return 'no';
	end;
	$$ language plpgsql;

create or replace function add_person(c_name varchar, c_client varchar,
c_tel varchar, c_mail varchar, c_info varchar, c_last_contact date) returns varchar as
	$$
	declare
		c varchar;
	begin
		select check_person(c_client, c_name) into c;
		if c = 'yes' then
			return 'no';
		end if;
		insert into persons (name, client, tel, mail, info, last_contact) values
			(c_name, c_client, c_tel, c_mail, c_info, c_last_contact);
		return 'ok';
		exception
			when others then
    			return 'no';
	end;
	$$ language plpgsql;

reate or replace function update_client(c_name varchar, c_category varchar, c_tag varchar, c_city varchar, c_address varchar,
c_tel varchar, c_mail varchar, c_site varchar, c_inst varchar, c_vk varchar, c_last_contact date) returns varchar as
	$$
	declare
		c varchar;
	begin
		select check_client(c_name) into c;
		if c = 'no' then
			return 'no';
		end if;

		update clients set category = c_category, tag = c_tag, city = c_city, address = c_address, 
			tel = c_tel, mail = c_mail, site = c_site, inst = c_inst, vk = c_vk, last_contact = c_last_contact 
			where name = c_name;
		return 'ok';
		exception
			when others then
    			return 'no';
	end;
	$$ language plpgsql;

create or replace function update_person(c_name varchar, c_client varchar,
c_tel varchar, c_mail varchar, c_info varchar, c_last_contact date) returns varchar as
	$$
	declare
		c varchar;
	begin
		select check_person(c_client, c_name) into c;
		if c = 'no' then
			return 'no';
		end if;
		update persons set tel = c_tel, mail = c_mail, info = c_info, last_contact = c_last_contact
			where name = c_name and client = c_client;
		return 'ok';
		exception
			when others then
    			return 'no';
	end;
	$$ language plpgsql;

create or replace function add_ord_main_info(o_id int, o_client varchar, o_person varchar, o_pay_date timestamp) returns varchar as
	$$
	declare
		c varchar;
		o_give_date timestamp;
	begin
		select check_id(o_id) into c;
		if c = 'ok' then
			return 'no';
		end if;

		if c = 'no' then
			o_give_date = o_pay_date + interval '20 days';
			insert into orders (id, client, person, state, pay_date, give_date) values
				(o_id, o_client, o_person, 'Оплачен', o_pay_date::date, o_give_date::date);
			return 'ok';
		end if;
		exception
			when others then
    			return 'no';
	end;
	$$ language plpgsql;

create or replace function add_ord_content(o_id int, o_color varchar, o_type varchar, o_shape varchar,
o_edge varchar, o_msq numeric, o_items int) returns varchar as
	$$
	declare
	begin
		insert into ord_content (id, color, type, shape, edge, msq, items) values
			(o_id, o_color, o_type, o_shape, o_edge, o_msq, o_items);
		return 'ok';
		exception
			when others then
    			return 'no';
	end;
	$$ language plpgsql;

create or replace function add_ord_money(o_id int, o_price numeric, o_prepay numeric, o_debt numeric, o_type varchar,
o_doc varchar, o_doc_date date, o_delivery numeric) returns varchar as
	$$
	declare
	begin
		insert into ord_money (id, price, prepay, debt, type, doc, doc_date, delivery) values
			(o_id, o_price, o_prepay, o_debt, o_type, o_doc, o_doc_date, o_delivery);
		return 'ok';
		exception
			when others then
    			return 'no';
	end;
	$$ language plpgsql;

create or replace function get_status(o_id int) returns varchar as 
	$$
	declare
		status varchar;
		c varchar;
	begin
		select check_id(o_id) into c;
		if c = 'no' then
			return 'no';
		end if;
		select state from orders where id = o_id into status;
		return status;
	end;
	$$ language plpgsql;

create or replace function get_status_date(o_id int, o_status varchar) returns date as
	$$
	declare
		res_date date;
	begin
		if o_status = 'Оплачен' then
			select pay_date from orders where id = o_id into res_date;
		end if;

		if o_status = 'Заказан' then
			select request_date from orders where id = o_id into res_date;
		end if;

		if o_status = 'На складе' then
			select recieve_date from orders where id = o_id into res_date;
		end if;

		if o_status = 'Выдан' then
			select give_date from orders where id = o_id into res_date;
		end if;

		return res_date;
	end;
	$$ language plpgsql;

create or replace function update_status(o_id int, new_status varchar, new_date date) returns varchar as
	$$
	declare
	begin
		if new_status = 'Оплачен' then
			update orders set state = new_status, pay_date = new_date where id = o_id;
			return 'ok';
		end if;

		if new_status = 'Заказан' then
			update orders set state = new_status, request_date = new_date where id = o_id;
			return 'ok';
		end if;

		if new_status = 'На складе' then
			update orders set state = new_status, recieve_date = new_date where id = o_id;
			return 'ok';
		end if;

		if new_status = 'Выдан' then
			update orders set state = new_status, give_date = new_date where id = o_id;
			return 'ok';
		end if;

		return 'no';
		exception
			when others then
    			return 'no';
	end;
	$$ language plpgsql;

create or replace function add_prepay(o_id int, o_prepay numeric) returns varchar as
	$$
	declare
	begin
		update ord_money set prepay = prepay + o_prepay, debt = debt - o_prepay where id = o_id;
		return 'ok';
		exception
			when others then
    			return 'no';
	end;
	$$ language plpgsql;

create or replace function set_bill(o_id int, o_bill numeric) returns varchar as
	$$
	declare
	begin
		update ord_money set bill = o_bill, profit = (price - o_bill - extra - delivery) where id = o_id;
		return 'ok';
		exception
			when others then
    			return 'no';
	end;
	$$ language plpgsql;

create or replace function add_extra(o_id int, o_extra numeric) returns varchar as
	$$
	declare
	begin
		update ord_money set extra = extra + o_extra, profit = profit - o_extra where id = o_id;
		return 'ok';
		exception
			when others then
    			return 'no';
	end;
	$$ language plpgsql;

