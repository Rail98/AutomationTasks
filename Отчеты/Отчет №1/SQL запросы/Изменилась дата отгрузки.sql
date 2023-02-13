DECLARE @Date1 DATE;
DECLARE @Date2 DATE;
--SET @Date1 = GETDATE()-1;			-- Вчера
--SET @Date2 = GETDATE()-2;			-- Позавчера

SET @Date1 = case when DATEPART(dw,GETDATE()-1) not in (1,7) then GETDATE()-1 else GETDATE()-3 end;			-- минус 1 рабочий день
SET @Date2 = case when DATEPART(dw,GETDATE()-1) not in (1,7) then GETDATE()-2 else GETDATE()-4 end;			-- минус 2 рабочих дня 



set dateformat dmy;	--- меняем ежедневно 2 даты

---- date_transp - дата прихода транспорта
---- date - дата отгрузки
---- date_change - дата изменения заявки (не нужна нам)
---- date_imp - дата создания заявки
----select * from order_status
--select * from co_material_attr_prod p where p.shipping_id in ('31-6','150-6')

--select distinct zh.zaivka_id, zh.transp_id from zakaz_hat zh where zh.date between '01.07.2016' and '31.07.2016'
--select * from zaivka
--select * from payment2

select list.pos, list.region_name, list.saleman_id, list.saleman, list.adress_d, list.address, list.date_imp, list.date_otgr_pl as date_otgr_pl1, list.date_otgr_fact, list.num, list.id, list.st, list.StatusName, list.kg
	, list.date_otgr_pl as date_otgr_pl2
	, pl_vchera
	, pl_pozavchera
	, isnull(pl_pozavchera,pl_vchera) pl_ish	
	
from
	(select r.pos, r.region_name, zh.saleman_id, c.name saleman, replace(zh.adress_d, '"', '')adress_d, cac.address
		, isnull(convert(char,zh.date_imp,104),'') date_imp, isnull(convert(char,zh.date,104),'') date_otgr_pl, isnull(convert(char,zh.date_otgr,104),'') date_otgr_fact
		, zh.num, zh.id, zh.st, os.StatusName
		, sum(zt.mass) kg
	from zakaz_hat zh
		inner join zakaz_tab zt on zt.id=zh.id
		inner join co_material_attr_prod p on p.material_id=zt.material_id
		inner join co_material cm on cm.id=p.material_id
		inner join co_contractor_attr_customer cac on cac.distr_id=zh.saleman_id
		inner join co_contractor c on c.id=cac.contractor_id
		inner join region2 r on r.region_id2=cac.region_id2
		left join order_status os on zh.st=os.id
	where c.factory_id=1 and c.del=0 and cm.factory_id=1 and cm.del=0
	and zh.date_imp>='01.01.2020'
	and zh.warehouse_id in (16952,16950,18534,18784,17831)
	and zh.st in (4,5,6,7,8)
	group by r.pos, r.region_name, zh.saleman_id, c.name, replace(zh.adress_d, '"', ''), cac.address
		, isnull(convert(char,zh.date_imp,104),''), isnull(convert(char,zh.date,104),''), isnull(convert(char,zh.date_otgr,104),'')
		, zh.num, zh.id, zh.st, os.StatusName) list

left join
(select zhh.id, isnull(convert(char,zhh.date,104),'') pl_vchera
from analitic.dbo.registry_zakaz_hat_history zhh
	inner join analitic.dbo.registry_zakaz rz on zhh.registry_zakaz_id=rz.id
where rz.type_id=2 and rz.date=@Date1									---минус 1 рабочий день (не зависит от итогов)
and zhh.st in (4,5,6,7,8)
and zhh.warehouse_id in (16952,16950,18534,18784,17831)) pl_1 on list.id=pl_1.id

left join
(select zhh.id, isnull(convert(char,zhh.date,104),'') pl_pozavchera
from analitic.dbo.registry_zakaz_hat_history zhh
	inner join analitic.dbo.registry_zakaz rz on zhh.registry_zakaz_id=rz.id
where rz.type_id=2 and rz.date=@Date2									----минус 2 рабочих дня (не зависит от итогов)
and zhh.st in (4,5,6,7,8)
and zhh.warehouse_id in (16952,16950,18534,18784,17831)) pl_2 on list.id=pl_2.id

where list.date_otgr_pl<>isnull(pl_pozavchera,pl_vchera)

order by list.pos, list.region_name, list.saleman_id, list.saleman, list.adress_d, list.address, list.date_imp, list.date_otgr_pl, list.date_otgr_fact, list.num, list.id, list.st, list.StatusName





