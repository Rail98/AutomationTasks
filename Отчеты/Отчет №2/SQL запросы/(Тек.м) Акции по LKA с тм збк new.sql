-- Изменять для каждого месяца в начале (в 11 местах)
-- Current Year/Month - заданный для выгрузки Год/Месяц 

/*DECLARE @currentYear	int = 2020;		-- текущий год
DECLARE @currentMonth	int = 05;		-- текущий месяц
DECLARE @previousYear	int = 2020;		-- год предыдущего месяца
DECLARE @previousMonth	int = 04;		-- предыдущий месяц
DECLARE @nextYear		int = 2020;		-- год следующего месяца

DECLARE @nextMonth		int = 06;		-- следующий месяц*/
/*DECLARE @firstDateOfCurrentMonth		date = '01.04.2020';
DECLARE @lastDateOfCurrentMonth			date = '30.04.2020';
DECLARE @firstDateOfPreviousMonth		date = '01.03.2020';
DECLARE @lastDateOfPreviousMonth		date = '31.03.2020';
DECLARE @firstDateOfMonthTwoMonthAgo	date = '01.02.2020';*/

declare
@firstDateOfCurrentMonth		date,
@lastDateOfCurrentMonth			date,
@firstDateOfPreviousMonth		date,
@lastDateOfPreviousMonth		date,
@firstDateOfMonthTwoMonthAgo    date,
@ffirstDateOfCurrentMonth1      date
 
set @firstDateOfCurrentMonth = dateadd(m,datediff(m,0,GetDate()),0);
set @lastDateOfCurrentMonth	 = dateadd(d, -1, dateadd(month,1, @firstDateOfCurrentMonth ));
SET @firstDateOfPreviousMonth = dateadd(m, DATEDIFF(m, 0,GetDate())-1, 0)
SET @lastDateOfPreviousMonth =  dateadd(d, -1, dateadd(month,1, @firstDateOfPreviousMonth ));
set @firstDateOfMonthTwoMonthAgo = dateadd(m, DATEDIFF(m, 0,GetDate())-2, 0)

--set @firstDateOfMonthTwoMonthAgo  = dateadd(m, 0, dateadd(d,1,@lastDateOfCurrentMonth ));-- следующая дата месяца


DECLARE @currentYear int = Year(@firstDateOfCurrentMonth);              -- текущий год
DECLARE @currentMonth int = Month(@firstDateOfCurrentMonth);            -- текущий месяц
DECLARE @previousYear	int = Year(@firstDateOfMonthTwoMonthAgo);		-- год предыдущего месяца
DECLARE @previousMonth	int =  @currentMonth -1;		                -- предыдущий месяц
DECLARE @nextYear		int = Year(@ffirstDateOfCurrentMonth1);		    -- год следующего месяца
DECLARE @nextMonth		int =  Month(@ffirstDateOfCurrentMonth1);		-- следующий месяц


--select distinct @currentMonth as 'Текущая дата' from co_material -- проверка даты

SET DATEFORMAT DMY;
SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;

select isnull(actio.pos,0) pos, isnull(actio.region_name,'')region_name, isnull(actio.network_id,'')network_id, isnull(actio.legal_name,'')legal_name, isnull(actio.network,'')network
, actio.typ
, case when actio.action_name is not NULL  then ROW_NUMBER () OVER ( partition by actio.network_id ORDER BY actio.action_name desc) else 0 end as row
, isnull(actio.TG,'')TG
, isnull(actio.TG_name,'')TG_name
, isnull(actio.priznak,'') priznak

, '=СЦЕПИТЬ(RC[-10];"S";RC[-8];"S";RC[-4])' kod1
, '=СЦЕПИТЬ(RC[-11];"S";RC[-9];"d";СУММ(RC[4]:RC[7]);"s";RC[8];"A";RC[1])' kod2 

, replace (isnull(actio.action_name,''),'"','') action_name
, actio.promo
--,actio.subaction_id
--,actio.action_id
---------------------------------------------------------------------------------------------------------
, isnull(convert(char,case when datediff(dd, actio.shipment_start, dateadd(dd, -day(dateadd(mm, 1, actio.shipment_start)), dateadd(mm, 1, actio.shipment_start))) + datediff(dd, DATEADD(dd, 1-day(actio.shipment_end), actio.shipment_end), actio.shipment_end) + 1 < datediff(dd, actio.shipment_start, actio.shipment_end) then dateadd(mm, -1, DATEADD(dd, 1-day(actio.shipment_end), actio.shipment_end))
							when datediff(dd, actio.shipment_start, dateadd(dd, -day(dateadd(mm, 1, actio.shipment_start)), dateadd(mm, 1, actio.shipment_start)))>datediff(dd, DATEADD(dd, 1-day(actio.shipment_end), actio.shipment_end), actio.shipment_end) then DATEADD(dd, 1-day(actio.shipment_start), actio.shipment_start)
							else DATEADD(dd, 1-day(actio.shipment_end), actio.shipment_end) end,104),'') month_name
---------------------------------------------------------------------------------------------------------
,isnull(convert(char,actio.shipment_start,104),'')shipment_start, isnull(convert(char,actio.shipment_end,104),'')shipment_end, isnull(convert(char,actio.holding_start,104),'')holding_start, isnull(convert(char,actio.holding_end,104),'')holding_end
, isnull(actio.status_id,'')status_id, isnull(actio.status,'')status , case when status_id in (4) then actio.status_comment end status_comment 

,isnull(convert(char,actio.filling_date,104),'')filling_date
, isnull(sum(actio.agreed_volume),0)agreed_volume
, isnull(actio.pt_name,'')pt_name
,isnull(sum(actio.cost),0)cost, isnull(sum(actio.amount),0)amount
, isnull(sum(actio.sht_pred),0)sht_pred, isnull(sum(actio.kg_pred),0)kg_pred, isnull(sum(actio.rub_prod_pred),0)rub_prod_pred, isnull(sum(actio.rub_zak_pred),0)rub_zak_pred
, isnull(sum(actio.sht),0) sht, isnull(sum(actio.kg),0)kg, isnull(sum(actio.rub_prod),0)rub_prod, isnull(sum(actio.rub_zak),0)rub_zak
, isnull(actio.status_comment,'')status_comment 

from
(select isnull(act.pos,'') pos, isnull(act.region_name,'')region_name, isnull(act.network_id,'')network_id, isnull(act.legal_name,'')legal_name, isnull(act.network,'')network
, act.typ
,act.promo,act.subaction_id
,act.action_id
, isnull(act.action_name,'')action_name,isnull(convert(char,act.shipment_start,104),'')shipment_start, isnull(convert(char,act.shipment_end,104),'')shipment_end, isnull(convert(char,act.holding_start,104),'')holding_start, isnull(convert(char,act.holding_end,104),'')holding_end
, isnull(act.status_id,'')status_id, isnull(act.status,'')status , isnull(act.status_comment,'')status_comment 
, isnull(act.filling_date,'')filling_date
, isnull(sum(act.agreed_volume),0)agreed_volume
, isnull(act.pt_name,'')pt_name,isnull(sum(act.cost),0)cost, isnull(sum(act.amount),0)amount
, isnull(sum(act.sht),0) sht, isnull(sum(act.kg),0)kg, isnull(sum(act.rub_prod),0)rub_prod, isnull(sum(act.rub_zak),0)rub_zak
, isnull(sum(act.sht_pred),0)sht_pred, isnull(sum(act.kg_pred),0)kg_pred, isnull(sum(act.rub_prod_pred),0)rub_prod_pred, isnull(sum(act.rub_zak_pred),0)rub_zak_pred
, isnull(act.TG,'')TG
, isnull(act.TG_name,'')TG_name
, isnull(act.priznak,'') priznak

from

(select isnull(ac.pos,'') pos, isnull(ac.region_name,'')region_name, isnull(ac.network_id,'')network_id, isnull(ac.legal_name,'')legal_name, isnull(ac.network,'')network
, ac.typ
,ac.promo
,ac.subaction_id ,ac.action_id
, isnull(ac.action_name,'')action_name
, isnull(convert(char,ac.shipment_start,104),'')shipment_start, isnull(convert(char,ac.shipment_end,104),'')shipment_end, isnull(convert(char,ac.holding_start,104),'')holding_start, isnull(convert(char,ac.holding_end,104),'')holding_end
, isnull(ac.class_id,'')class_id, isnull(ac.class,'')class, isnull(ac.format_id,'')format_id, isnull(ac.format,'')format
, isnull(ac.bonus_rf,0)bonus_rf, isnull(ac.discount,0)discount, isnull(ac.boxes,0)boxes, isnull(ac.one_to_one,0)one_to_one
, isnull(ac.order_amount,0)order_amount, isnull(ac.order_sum,0)order_sum, isnull(ac.compensation_type,'')compensation_type
, isnull(ac.status_id,'')status_id, isnull(ac.status,'')status
, isnull(ac.comment,'')status_comment 
, isnull(ac.filling_date,'')filling_date
, isnull(ac.agreed_volume,0) agreed_volume
, isnull(ac.pt_id,'')pt_id, isnull(ac.pt_name,'')pt_name, isnull(ac.cost,0)cost, isnull(ac.amount,0)amount
, isnull(sum(sale.sht),0) sht, isnull(sum(sale.kg),0)kg, isnull(sum(sale.rub_prod),0)rub_prod, isnull(sum(sale.rub_zak),0)rub_zak
, isnull(sum(sale_pred.sht_pred),0)sht_pred, isnull(sum(sale_pred.kg_pred),0)kg_pred, isnull(sum(sale_pred.rub_prod_pred),0)rub_prod_pred, isnull(sum(sale_pred.rub_zak_pred),0)rub_zak_pred
, isnull(ac.cg,'') TG
, isnull(ac.tg,'') TG_name
, isnull(ac.priznak,'') priznak



from 
(select r.pos, r.region_name, sn5.id network_id, sn5.legal_name, sn5.name network
			,case when snv3.network_attr_id=31 then 'DLKA'else  'LKA'  end typ
			, sub.subaction_id
			, a.action_id
			, ISNULL (als.large_scale_share,'')pr
			,case when als.large_scale_share=1 then 'Акция масштабное промо' else  ''  end promo
			, a.action_name, an.shipment_start, an.shipment_end, an.holding_start, an.holding_end, a.class_id, ac.name class, a.format_id, af.name format
			, sub.bonus_rf, sub.discount, sub.boxes, sub.one_to_one, sub.order_amount, sub.order_sum, sub.compensation_type
			, case when an.status_id=1 then 3 when an.status_id=3 then 1  else an.status_id end status_id
			, case when an.status_id=1 then 'Черновик' else case when an.status_id=2 then 'На согласовании' else case when an.status_id=3 then 'Согласована' else case when an.status_id=4 then 'Отклонена' end end end end status
			, isnull(sum(snaa.agreed_volume),0) agreed_volume
			, isnull(prom_type.id,'') pt_id, isnull(prom_type.name,'')pt_name
			--, isnull(sum(prom_type.cost),0) pt_cost, isnull(sum(prom_type.amount),0) pt_amount
			, prom_type.cost, prom_type.amount
			, agr.group_id, agr.group_name
			, prod.material_id, cm.name, cmp.cg
			, case when cmp.brand=635 and cmp.cg = 16 then 'ТМ ЗБК' else t.tg end tg -----------------------------------ЗБК
			--, t.tg------------------------------------------------------------------------------------------------------
			, an.filling_date
			, apsh.comment
			, case when cmp.warehouse_id in (19533) then 'Аромат' else 'НК' end priznak
			
	from action a
		inner join action_format af on a.format_id=af.id
		inner join action_classification ac on a.class_id=ac.id
		inner join action_subactions sub on a.action_id=sub.action_id
		inner join action_groups agr on sub.subaction_id=agr.subaction_id
		inner join action_prodlist prod on agr.group_id=prod.group_id
		inner join co_material cm on cm.id=prod.material_id
		inner join co_material_attr_prod cmp on cm.id=cmp.material_id
		inner join tg t on cmp.cg=t.tg_id
		inner join action_additional_params an on sub.subaction_id=an.subaction_id
		inner join sam_network sn on sn.id=an.network_id and sn.factory_id=1 and sn.del=0
		inner join sam_network_attr_value snv on snv.network_id=sn.id and snv.network_attr_id=2
		inner join (select sn.id network_id, sn.name, isnull(snv.network_attr_id,32) network_attr_id
					from sam_network sn
						left join sam_network_attr_value snv on sn.id=snv.network_id and snv.network_attr_id in (31,32)
					where sn.factory_id=1 and sn.del=0
					group by sn.id, sn.name, snv.network_attr_id) snv2 on snv2.network_id=sn.id and snv2.network_attr_id=32 ---Непрямые сети
		left join dbo.action_large_shares als on sn.id=als.network_id and als.subaction_id=sub.subaction_id
		left join sam_network sn2 on sn2.id = isnull(sn.parent, sn.id)
		left join sam_network sn3 on sn3.id = isnull(sn2.parent, sn2.id)
		left join sam_network sn4 on sn4.id = isnull(sn3.parent, sn3.id)
		left join sam_network sn5 on sn5.id = isnull(sn4.parent, sn4.id)
		left join sam_network_region snr on snr.network_id=sn.id and snr.region_id<>12
		left join region2 r on r.region_id2=snr.region_id
		left join sam_network_action_agreement snaa on an.subaction_id=snaa.subaction_id and an.network_id=snaa.network_id
		left join action_params_status_history apsh  on apsh.action_params_id = an.id 
		inner join (select MAX(date_change) date_change, aap.subaction_id
		                   from  action_params_status_history apsh inner join action_additional_params aap on aap.id = apsh.action_params_id
		                     where apsh.date_change  > '01.01.2019'
		                     group by aap.subaction_id) b on b.subaction_id = an.subaction_id and b.date_change  = apsh.date_change
		                     
		left join (SELECT action_params_id, pt.id, pt.name ,cost, amount
				  FROM action_promotion ap
				inner join co_promotion_types pt on pt.id = ap.promotion_id) prom_type on prom_type.action_params_id = an.id

		left join sam_network_attr_value snv3 on snv3.network_id=sn5.id and snv3.network_attr_id in (31,32)
	where a.factory_id=1
	
	and ((year(an.shipment_start)=@currentYear and month(an.shipment_start)=@currentMonth)
	or (year(an.shipment_end)=@currentYear and month(an.shipment_end)=@currentMonth) 
	or (year(an.shipment_start)=@previousYear and year(an.shipment_end)=@nextYear 
	and month(an.shipment_start)=@previousMonth and month(an.shipment_end)=@nextMonth
	and month(an.shipment_start)<>@currentMonth and month(an.shipment_end)<>@currentMonth))
		
	and a.class_id in (4,6,8)-----акции со скидкой (8) -- бонус и примотка 1+1 (4,6)
	and an.status_id not in (1)
		
	group by r.pos, r.region_name, sn5.id, sn5.legal_name, sn5.name
			, sub.subaction_id,als.large_scale_share, a.action_id
			, a.action_name, an.shipment_start, an.shipment_end, an.holding_start, an.holding_end, a.class_id, ac.name, a.format_id, af.name
			, sub.bonus_rf, sub.discount, sub.boxes, sub.one_to_one, sub.order_amount, sub.order_sum, sub.compensation_type
			, case when an.status_id=1 then 3 when an.status_id=3 then 1  else an.status_id end 
			, case when an.status_id=1 then 'Черновик' else case when an.status_id=2 then 'На согласовании' else case when an.status_id=3 then 'Согласована' else case when an.status_id=4 then 'Отклонена' end end end end
			, prom_type.id, prom_type.name, prom_type.cost, prom_type.amount
			, agr.group_id, agr.group_name
			, prod.material_id, cm.name, cmp.cg, t.tg,an.filling_date, snv.network_attr_id
			,case when snv3.network_attr_id=31 then 'DLKA'else  'LKA'  end
			, apsh.comment
			, cmp.brand--------------
			, case when cmp.warehouse_id in (19533) then 'Аромат' else 'НК' end
			) ac
			
			left join
	------Продажи за отчётный период
	(select sn5.id tnet_id, mt.material_id,prod.group_id
		, isnull(sum(mt.VAMOUNT),0) sht
		, isnull(sum(mt.VAMOUNT*unit_mass/1000),0) kg
		, isnull(sum(convert(float,mt.VSUM_RUB)),0) rub_prod
		, isnull(sum(case when md.contractor_id=15342 then mt.VSUM_RUB else mt.VAMOUNT*p.unit_mass/1000*pd.cost_nds end),0) rub_zak
	from MT_DOCS_db md
		inner join mt_details_db mt on mt.RDOC_ID=md.DOC_ID
		inner join co_material cm on cm.id=mt.material_id
		inner join co_material_attr_prod_db p on p.material_id=mt.material_id
		inner join client_card cc on cc.client_id=md.nid
		inner join co_contractor c on c.id=md.contractor_id
		inner join co_contractor_attr_customer cac on cac.contractor_id=c.id
		inner join action_prodlist prod on prod.material_id=mt.material_id
		inner join action_groups agr on prod.group_id=agr.group_id
		inner join action_subactions sub on agr.subaction_id=sub.subaction_id
		inner join action a on sub.action_id=a.action_id and a.factory_id=1
		inner join action_additional_params an on sub.subaction_id=an.subaction_id and cc.tnet_id=an.network_id 
		
	and ((year(an.shipment_start)=@currentYear and month(an.shipment_start)=@currentMonth)
	or (year(an.shipment_end)=@currentYear and month(an.shipment_end)=@currentMonth) 
	or (year(an.shipment_start)=@previousYear and year(an.shipment_end)=@nextYear 
	and month(an.shipment_start)=@previousMonth and month(an.shipment_end)=@nextMonth
	and month(an.shipment_start)<>@currentMonth and month(an.shipment_end)<>@currentMonth))
		
		left join price_distr_db pd on pd.contractor_id=md.contractor_id and pd.material_id=mt.material_id and pd.date=DATEADD(dd,-day(dateadd(mm,1,md.vdate)),dateadd(mm,1,md.vdate))
		left join sam_network sn on sn.id = cc.tnet_id
		left join sam_network sn2 on sn2.id = isnull(sn.parent, sn.id)
		left join sam_network sn3 on sn3.id = isnull(sn2.parent, sn2.id)
		left join sam_network sn4 on sn4.id = isnull(sn3.parent, sn3.id)
		left join sam_network sn5 on sn5.id = isnull(sn4.parent, sn4.id)
		
	where md.VDATE between @firstDateOfCurrentMonth AND @lastDateOfCurrentMonth and md.VDATE between an.shipment_start and an.shipment_end					--***************************************
	and ftype=0 and RSTATE=1
	and md.FDEL=0 and mt.FDEL=0
	and RDOC_TYPE in (2,5)
	and md.contractor_id not in (15342,4150)
	and cm.factory_id=1 and cm.del=0
	and c.factory_id=1 and c.del=0
	and cc.factory_id=1
	group by sn5.id, mt.material_id,prod.group_id) sale on ac.network_id=sale.tnet_id and ac.group_id=sale.group_id and ac.material_id=sale.material_id
	
			left join
	------Продажи за предыдущий период
	(select sn5.id tnet_id, mt.material_id,prod.group_id
		, isnull(sum(mt.VAMOUNT),0) sht_pred
		, isnull(sum(mt.VAMOUNT*unit_mass/1000),0) kg_pred
		, isnull(sum(convert(float,mt.VSUM_RUB)),0) rub_prod_pred
		, isnull(sum(case when md.contractor_id=15342 then mt.VSUM_RUB else mt.VAMOUNT*p.unit_mass/1000*pd.cost_nds end),0) rub_zak_pred
	from MT_DOCS_db md
		inner join mt_details_db mt on mt.RDOC_ID=md.DOC_ID
		inner join co_material cm on cm.id=mt.material_id
		inner join co_material_attr_prod_db p on p.material_id=mt.material_id
		inner join client_card cc on cc.client_id=md.nid
		inner join co_contractor c on c.id=md.contractor_id
		inner join co_contractor_attr_customer cac on cac.contractor_id=c.id
		inner join action_prodlist prod on prod.material_id=mt.material_id
		inner join action_groups agr on prod.group_id=agr.group_id
		inner join action_subactions sub on agr.subaction_id=sub.subaction_id
		inner join action a on sub.action_id=a.action_id and a.factory_id=1
		inner join action_additional_params an on sub.subaction_id=an.subaction_id and cc.tnet_id=an.network_id 
		
	and ((year(an.shipment_start)=@currentYear and month(an.shipment_start)=@currentMonth)
	or (year(an.shipment_end)=@currentYear and month(an.shipment_end)=@currentMonth) 
	or (year(an.shipment_start)=@previousYear and year(an.shipment_end)=@nextYear 
	and month(an.shipment_start)=@previousMonth and month(an.shipment_end)=@nextMonth
	and month(an.shipment_start)<>@currentMonth and month(an.shipment_end)<>@currentMonth))
		
		
		left join price_distr_db pd on pd.contractor_id=md.contractor_id and pd.material_id=mt.material_id and pd.date=DATEADD(dd,-day(dateadd(mm,1,md.vdate)),dateadd(mm,1,md.vdate))
		left join sam_network sn on sn.id = cc.tnet_id
		left join sam_network sn2 on sn2.id = isnull(sn.parent, sn.id)
		left join sam_network sn3 on sn3.id = isnull(sn2.parent, sn2.id)
		left join sam_network sn4 on sn4.id = isnull(sn3.parent, sn3.id)
		left join sam_network sn5 on sn5.id = isnull(sn4.parent, sn4.id)
	where md.VDATE between @firstDateOfPreviousMonth AND @lastDateOfPreviousMonth and md.VDATE between an.shipment_start and an.shipment_end					--***************************************
	and ftype=0 and RSTATE=1
	and md.FDEL=0 and mt.FDEL=0
	and RDOC_TYPE in (2,5)
	and md.contractor_id not in (15342,4150)
	and cm.factory_id=1 and cm.del=0
	and c.factory_id=1 and c.del=0
	and cc.factory_id=1
	group by sn5.id, mt.material_id,prod.group_id) sale_pred on ac.network_id=sale_pred.tnet_id and ac.group_id=sale_pred.group_id and ac.material_id=sale_pred.material_id
	--where pos in (5)
	
	group by ac.pos, ac.region_name, ac.network_id, ac.legal_name,ac.network
							, ac.action_name,ac.shipment_start, ac.shipment_end, ac.holding_start, ac.holding_end, ac.class_id, ac.class, ac.format_id, ac.format
							, ac.bonus_rf, ac.discount, ac.boxes, ac.one_to_one, ac.order_amount, ac.order_sum, ac.compensation_type
							, ac.status_id, ac.status
							, ac.agreed_volume
							, ac.pt_id, ac.pt_name, ac.cost, ac.amount, ac.cg, ac.tg, ac.filling_date, ac.typ,ac.promo,ac.subaction_id,ac.action_id,ac.comment
							, ac.priznak) act
							
							group by act.pos, act.region_name, act.network_id, act.legal_name,act.network
							, act.action_name,act.shipment_start, act.shipment_end, act.holding_start, act.holding_end
							, act.status_id, act.status, act.pt_name, act.TG, act.TG_name, act.filling_date,act.typ,act.promo,act.subaction_id,act.action_id,act.status_comment, act.priznak) actio
	
	group by actio.pos, actio.region_name, actio.network_id, actio.legal_name,actio.network
							, actio.action_name,actio.shipment_start, actio.shipment_end, actio.holding_start, actio.holding_end,actio.TG, actio.TG_name, actio.filling_date, actio.priznak
							--,actio.agreed_volume
							/*, actio.class_id, actio.class, actio.format_id, actio.format
							, actio.bonus_rf, actio.discount, actio.boxes, actio.one_to_one, actio.compensation_type*/
							, actio.status_id, actio.status
							,case when status_id in (4) then actio.status_comment end
							, actio.typ,actio.promo,actio.subaction_id,actio.action_id
							--, actio.pt_id
							, actio.pt_name
							,status_comment	
	
	/*group by grouping sets ((actio.pos, actio.region_name, actio.network_id, actio.legal_name,actio.network
							, actio.action_name,actio.shipment_start, actio.shipment_end, actio.holding_start, actio.holding_end
							,actio.agreed_volume
							/*, actio.class_id, actio.class, actio.format_id, actio.format
							, actio.bonus_rf, actio.discount, actio.boxes, actio.one_to_one, actio.compensation_type*/
							, actio.status_id, actio.status
							--, actio.pt_id
							, actio.pt_name)
							, (actio.pos, actio.region_name, actio.network_id, actio.legal_name,actio.network),(actio.pos, actio.region_name),())*/
							
							
UNION 


----Сети без акций (активные за 3 месяца)
SELECT 
isnull(r.pos, 0)
,r.region_name
,sn5.id 
,sn5.legal_name
,sn5.name
,case when snv3.network_attr_id=31 then 'DLKA'else  'LKA'  end typ
, '', '', ''
,'Нет акций'
--, '', '', '', '', '', '', '', '', '', ''
--, '', '', ''
, '', '', '', '', '', '', '', ''
, '', '', '', '', '', '', '', '', '', '','','','','','','','',''
 FROM mt_docs md
join mt_details mt on md.DOC_ID = mt.RDOC_ID
join client_card c on c.client_id = md.NID
INNER JOIN sam_network sn ON sn.id = c.tnet_id
INNER JOIN sam_network sn2 ON sn2.id=isnull(sn.parent,sn.id)
INNER JOIN sam_network sn3 ON sn3.id=isnull(sn2.parent,sn2.id)
INNER JOIN sam_network sn4 ON sn4.id=isnull(sn3.parent,sn3.id)
INNER JOIN sam_network sn5 ON sn5.id=isnull(sn4.parent,sn4.id)
LEFT JOIN sam_network_region snr ON snr.network_id = sn.id AND snr.region_id<>12
LEFT JOIN region2 r ON r.region_id2 = snr.region_id
INNER JOIN sam_network_attr_value snav ON snav.network_id = sn.id
left join sam_network_attr_value snv3 on snv3.network_id=sn5.id and snv3.network_attr_id in (31,32)

WHERE md.VDATE between @firstDateOfMonthTwoMonthAgo and @lastDateOfCurrentMonth																					--***************************************
and md.RSTATE = 1
and md.RDOC_TYPE in (2,5)
and md.FDEL=0
and md.RSTATE=1
and mt.FDEL=0
and mt.FTYPE=0 
and sn.factory_id=1 
AND sn.del<>1 
AND sn.[active]=1 
and snav.network_attr_id=2
and isnull(snv3.network_attr_id,32)<>31
--and snr.region_id<>12
and sn5.id not in (select sn5.id
	from action a
		inner join action_format af on a.format_id=af.id
		inner join action_classification ac on a.class_id=ac.id
		inner join action_subactions sub on a.action_id=sub.action_id
		inner join action_groups agr on sub.subaction_id=agr.subaction_id
		inner join action_prodlist prod on agr.group_id=prod.group_id
		inner join co_material cm on cm.id=prod.material_id
		inner join action_additional_params an on sub.subaction_id=an.subaction_id
		inner join sam_network sn on sn.id=an.network_id and sn.factory_id=1 and sn.del=0
		inner join sam_network_attr_value snv on snv.network_id=sn.id and snv.network_attr_id=2
		inner join (select sn.id network_id, sn.name, isnull(snv.network_attr_id,32) network_attr_id
					from sam_network sn
						left join sam_network_attr_value snv on sn.id=snv.network_id and snv.network_attr_id in (31,32)
					where sn.factory_id=1 and sn.del=0
					group by sn.id, sn.name, snv.network_attr_id) snv2 on snv2.network_id=sn.id and snv2.network_attr_id=32 ---Непрямые сети
		left join sam_network sn2 on sn2.id = isnull(sn.parent, sn.id)
		left join sam_network sn3 on sn3.id = isnull(sn2.parent, sn2.id)
		left join sam_network sn4 on sn4.id = isnull(sn3.parent, sn3.id)
		left join sam_network sn5 on sn5.id = isnull(sn4.parent, sn4.id)
		left join sam_network_region snr on snr.network_id=sn.id and snr.region_id<>12
		left join region2 r on r.region_id2=snr.region_id
		left join sam_network_action_agreement snaa on an.subaction_id=snaa.subaction_id and an.network_id=snaa.network_id
		left join (SELECT action_params_id, pt.id, pt.name ,cost, amount
				  FROM action_promotion ap
				inner join co_promotion_types pt on pt.id = ap.promotion_id) prom_type on prom_type.action_params_id = an.id
	where a.factory_id=1
	
	and ((year(an.shipment_start)=@currentYear and month(an.shipment_start)=@currentMonth)
	or (year(an.shipment_end)=@currentYear and month(an.shipment_end)=@currentMonth) 
	or (year(an.shipment_start)=@previousYear and year(an.shipment_end)=@nextYear 
	and month(an.shipment_start)=@previousMonth and month(an.shipment_end)=@nextMonth
	and month(an.shipment_start)<>@currentMonth and month(an.shipment_end)<>@currentMonth))
		
	
	and a.class_id in (4,6,8)-----акции со скидкой (8) -- бонус и примотка 1+1 (4,6)
	and an.status_id not in (1)
	group by sn5.id)
GROUP BY r.pos
,r.region_name
,sn5.id 
,sn5.legal_name
,sn5.name	,case when snv3.network_attr_id=31 then 'DLKA'else  'LKA'  end				
							
	order by 1,3,9 desc