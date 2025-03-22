<?php

$skipped_columns = array("route_", "ctr_", "chm_store", "org_banner", "back_splash", "banner_", "road_", "interchange_", "logo_picture", "pk_"); // geo_

$fields_export_mapping = [
    ["fil", "wrk_time", 0, 1],
    ["wrk_time", "schedule", 1, 0],
    ["fil", "wrk_time_comment", 1, 3],
    ["rub3", "rub2", 0, 1],
    ["rub2", "rub1", 0, 1],
    ["rub1", "name", 1, 0],
    ["rub2", "name", 1, 0],
    ["rub3", "name", 1, 0],
    ["bld_purpose", "name", 1, 0],
    ["building", "purpose", 0, 2],
    ["bld_name", "name", 1, 0],
    ["building", "name", 0, 2],
    ["building", "post_index", 1, 3],
    ["map_to_building", "data", 0, 4],
    # payments
    ["fil_payment", "fil", 0, 1],
    ["fil_payment", "payment", 0, 2],
    ["payment_type", "name", 1, 0],

    ["fil_contact", "comment", 1, 3],
    ["address_elem", "map_oid", 0, 2],
    ["org", "id", 0, 2],
    ["org", "name", 1, 0],
    ["org_rub", "org", 0, 1],
    ["fil_contact", "type", 0, 2],
    ["fil_rub", "fil", 0, 1],
    ["fil_rub", "rub", 0, 2],
    ["address_elem", "building", 1, 0],
    ["city", "name", 1, 0],
    ["fil_contact", "comment", 1, 3],
    ["org_rub", "rub", 0, 2],

    ["address_elem", "street", 0, 1],
    ["street", "name", 1, 0],
    ["street", "city", 0, 1],

    ["fil_contact", "fil", 0, 1],
    ["fil_contact", "phone", 1, 0],
    ["fil_contact", "eaddr", 1, 0],
    ["fil_contact", "eaddr_name", 1, 3],

/*
    ["", "", 0, 1],
    ["", "", 0, 1],
*/
    ["fil_address", "fil", 0, 1],
    ["fil_address", "address", 0, 2],
    ["fil", "org", 0, 1],
];


$xlsx_export_cols = [
	"ID" => 20,
	"Название организации" => 40,
	"Населенный пункт" => 20,
	"Раздел" => 40,
	"Подраздел" => 40,
	"Рубрика" => 40,
	"Телефоны" => 30,
	"Факсы" => 30,
	"Email" => 20,
	"Сайт" => 20,
	"Адрес" => 30,
	"Почтовый индекс" => 10,
	"Типы платежей" => 20,
	"Время работы" => 34,
	"Собственное название строения" => 25,
	"Назначение строения" => 25,
	"Vkontakte" => 20,
	"Facebook " => 20,
	"Skype    " => 20,
	"Twitter  " => 20,
	"Instagram" => 20,
	"ICQ" => 20,
	"Jabber   " => 20,
];

?>