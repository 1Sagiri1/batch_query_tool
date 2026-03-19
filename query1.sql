select *
from
(SELECT o.waybill_num,
    o.source_num,
    w.station_name,
    p.package_no,
    lb.lading_bill_code,
    ROW_NUMBER() OVER (
      PARTITION BY o.waybill_num
      ORDER BY CASE WHEN lb.lading_bill_uid IS NOT NULL THEN 0 ELSE 1 END
    ) AS rn
FROM global_express_json_order o
left join work_station w on w.station_uid = o.station_uid
left JOIN global_express_package_rel pr on o.waybill_num = pr.waybill_num
left JOIN global_express_package p 
  ON p.package_uid = pr.package_uid
left JOIN global_express_lading_bill lb 
  ON lb.lading_bill_uid = p.lading_bill_uid
where o.waybill_num in ({})
) o
where o.rn = 1