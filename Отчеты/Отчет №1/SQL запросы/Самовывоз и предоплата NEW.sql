set dateformat dmy;


 
select r.pos, r.region_name, zh.saleman_id, c2.name saleman
	
	, replace(zh.adress_d, '"', '')adress_d, cac.address
	, zh.warehouse_id, c4.name warehouse
	, isnull(convert(char,zh.date_imp,104),'') date_imp, isnull(convert(char,zh.date,104),'') date_otgr_pl
	, zh.num

	, zh.id, zh.st, os.StatusName, case when zt.otgr=1 then 'отгружена' else '' end otgr
	
	,z.pallet palt
	, sum(zt.mass) kg 
	,p2.pay_name

, case when zh.transp_id='self' then 'самовывоз'  when zh.transp_id='selfwagon' then 'Вагон(самовывоз)' else zh.transp_id end transp 
, case when zh.transp_id in ('self','selfwagon') then 1 else 2 end transp_id

from zakaz_hat zh
	inner join zakaz_tab zt on zt.id=zh.id
	inner join co_material_attr_prod p on p.material_id=zt.material_id
	inner join co_material cm on cm.id=p.material_id
	left join analitic.dbo.tg_2 on tg_2.tg_id=p.cg
	inner join co_contractor_attr_customer cac on cac.distr_id=zh.consignee_id
	inner join co_contractor c on c.id=cac.contractor_id
	inner join region2 r on r.region_id2=cac.region_id2
	left join order_status os on zh.st=os.id
	left join co_contractor_attr_customer cac2 on cac2.distr_id=zh.saleman_id
	left join co_contractor c2 on c2.id=cac2.contractor_id
	left join co_contractor_attr_customer cac3 on cac3.distr_id=cac.buyer_id
	left join co_contractor c3 on c3.id=cac3.contractor_id
	left join co_contractor c4 on zh.warehouse_id=c4.id
	left join price_head ph on zh.price_id=ph.price_id and ph.factory_id=1 and ph.del=0
	left join price_body pb on pb.price_id=zh.price_id and pb.material_id=zt.material_id and pb.del=0
	left join payment2 p2 on p2.pay_id=zh.pay_id and p2.factory_id=1
	LEFT JOIN 
		(
		SELECT zt.id
		
		  , ceiling(SUM(kor/poddon_kor)) pallet
		FROM
		  zakaz_tab zt,
		  dbo.prod_pogruz pp,
		  dbo.pogruz pog, 
		  zakaz_hat zh 
		WHERE
		  
		  zt.material_id = pp.material_id AND
		
		  pp.pogruz_id = pog.pogruz_id AND
		  zt.kor <> 0 
		  
		  and zt.id=zh.id
		
		  and zh.zaivka_id=pp.zaivka_id
		  and zh.ver_id=pp.ver_id

		 group by zt.id
		
		)z ON z.id=zh.id 
	left join 
	(select tt.order_id, isnull(tt.transp_id,'')transp_id, isnull(tt.driver_info,'')driver_info, isnull(tt.number_auto,'')number_auto
   , isnull(convert(datetime,ttr.date_load_ready,104),'') 'Дата прибытия к воротам', isnull(convert(datetime,th.new_time,104),'') 'Дата заезда на комбинат', isnull(convert(datetime,ttr.date_shipped_fact,104),'') 'Дата выезда с комбината'
    from tc_transp tt
	left join tc_history th on th.transp_id=tt.transp_id and th.old_warehouse_id=999
	left join tc_trip ttr on ttr.ts_transp_id=tt.transp_id
	inner join (select tt1.order_id, max(th1.new_time)new_time
				from tc_transp tt1
				inner join tc_history th1 on tt1.transp_id=th1.transp_id and old_warehouse_id=999
				where
				tt1.warehouse_id<>-1
				group by tt1.order_id) m_tt on m_tt.order_id=tt.order_id and th.new_time=m_tt.new_time
    where tt.warehouse_id<>-1
    group by tt.order_id, isnull(tt.transp_id,''), isnull(tt.driver_info,''), isnull(tt.number_auto,'')
   , isnull(convert(datetime,ttr.date_load_ready,104),''), isnull(convert(datetime,th.new_time,104),''), isnull(convert(datetime,ttr.date_shipped_fact,104),''))tt on zh.id=tt.order_id
   
where c.factory_id=1 and c.del=0 and cm.factory_id=1 and cm.del=0

and zh.[date]>='01.08.2021'	
and zh.warehouse_id in (16952,16950,18534,18784,17831)
and zh.st in (4,5,7,8) and ((zt.otgr=0) or (zt.otgr is NULL)) --and zt.tip=0
and (zh.transp_id in ('self','selfwagon') or zh.pay_id=1)

and r.region_id2 not in (14)
group by r.pos, r.region_name, zh.saleman_id, c2.name
	, replace(zh.adress_d, '"', ''), cac.address
	, zh.warehouse_id, c4.name
	, isnull(convert(char,zh.date_imp,104),''), isnull(convert(char,zh.date,104),'')
	, zh.num

	, zh.id, zh.st, os.StatusName, case when zt.otgr=1 then 'отгружена' else '' end

	,z.pallet

	,p2.pay_name

, zh.transp_id

order by case when zh.transp_id in ('self','selfwagon') then 1 else 2 end, zh.transp_id, r.pos, zh.saleman_id, zh.num, zh.id





