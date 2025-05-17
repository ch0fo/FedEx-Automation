SELECT DISTINCT AL1.shp_trk_nbr FROM SCAN_PROD_VIEW_DB.fx_package_event_history_all AL1 WHERE (

{awbs}

 AND AL1.scan_type_cd IN ('11', '19', '20', '30', '31', '33', '42', '79') AND AL1.loc_arpt_cd LIKE 'Y%' AND AL1.evnt_crt_tmstp>TIMESTAMP '2022-03-01 00:00:00.000')
;