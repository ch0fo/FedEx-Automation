with broker as
(SELECT
    classify.shipment.local_shipment_oid_nbr             AS oid,
    classify.shipment.awb_nbr,
    classify.shipment.transaction_nbr                    AS tn,
    classify.shipment.duty_bill_to_cd,
    classify.shipment.last_modified_tmstp,
    classify.shipment.origin_loc_cntry_cd                AS coe,
    classify.shipment.piece_qty,
    classify.shipment.rod_flg,
    classify.shipment.ship_dt,
    classify.shipment.shipment_desc,
    classify.shipment.final_import_clearance_loc_cd      AS clr_loc,
    classify.shipment.eci_flg,
    classify.shipment.LOCAL_CUSTOMS_VALUE_AMT                        as cad_val,
    classify.shipment.entry_dt,
    classify.shipment_party.broker_id_cd                 AS brokr_id
FROM
    classify.shipment
    INNER JOIN classify.shipment_party ON classify.shipment.local_shipment_oid_nbr = classify.shipment_party.local_shipment_oid_nbr
    INNER JOIN CLASSIFY.SHIPMENT_PROCESS_CONTROL ON CLASSIFY.shipment.LOCAL_SHIPMENT_OID_NBR = CLASSIFY.SHIPMENT_PROCESS_CONTROL.LOCAL_SHIPMENT_OID_NBR
WHERE
Classify.SHIPMENT_PROCESS_CONTROL.PROCESS_CONTROL_TYPE_DESC = 'ENTRYS'
    and classify.shipment_party.broker_id_cd  in ('FEC','FON','FEX')
    and classify.shipment.origin_loc_cntry_cd not in ('US', 'MX','PR')
    and classify.shipment.manuf_orig_cntry_cd not in ('US')
    AND classify.shipment.entry_dt between '{start_date}' and '{end_date}'
and CLASSIFY.SHIPMENT_PROCESS_CONTROL.PROCESS_CONTROL_DATA_DESC in ('LVS')
    AND classify.shipment_party.shipment_party_type_cd = 'B' 
                and classify.shipment.LOCAL_CUSTOMS_VALUE_AMT <= '20'), impInfo as 
(SELECT
    broker.*,
    classify.shipment_party.Company_NM      as iCompany,
    classify.shipment_party.Contact_NM      as iContact,
                                classify.shipment_party.customer_acct_nbr as duty_bill_to_acct_nbr
FROM
    broker
    INNER JOIN classify.shipment_party ON broker.oid = classify.shipment_party.local_shipment_oid_nbr
WHERE
    classify.shipment_party.shipment_party_type_cd = 'I'), shipInfo as

(SELECT
    impInfo.*,
    classify.shipment_party.Company_NM      as sCompany,
    classify.shipment_party.Contact_NM      as sContact

FROM
    impInfo
    INNER JOIN classify.shipment_party ON impInfo.oid = classify.shipment_party.local_shipment_oid_nbr
WHERE
    classify.shipment_party.shipment_party_type_cd = 'S'), CompleteInfo as
(SELECT
    shipInfo.*,
    classify.shipment_party.Company_NM      as cCompany,
    classify.shipment_party.Contact_NM      as cContact         
FROM
    shipInfo
INNER JOIN classify.shipment_party ON shipInfo.oid = classify.shipment_party.local_shipment_oid_nbr
WHERE
    classify.shipment_party.shipment_party_type_cd = 'C')

    
select distinct /*CSV*/
    CompleteInfo.awb_nbr,
    trim(CompleteInfo.iCompany || ' ' || CompleteInfo.iContact) as importerNme,
    CompleteInfo.entry_dt,
    CompleteInfo.duty_bill_to_acct_nbr as billAcc,
    CompleteInfo.cad_val,
    CompleteInfo.brokr_id,
    CompleteInfo.duty_bill_to_cd,
CompleteInfo.clr_loc,
CompleteInfo.piece_qty as pcs,
    CompleteInfo.shipment_desc,
    CompleteInfo.cCompany,
    CompleteInfo.cContact,
    CompleteInfo.coe,
    CompleteInfo.tn,
    trim(CompleteInfo.sCompany || ' ' || CompleteInfo.sContact) as shipperNme

    from CompleteInfo
LEFT OUTER JOIN classify.dt_shipment_header ON (CompleteInfo.oid=classify.dt_shipment_header.local_shipment_oid_nbr) 
where classify.dt_shipment_header.local_shipment_oid_nbr is null