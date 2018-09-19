update ice set ice.fsourceEntryID=ss.源单分录,ice.fsourcebillNo=ss.源单单号,ice.FSourceInterId=ss.源单内码,ice.FSourceTranType=21
from icsaleentry ice
inner join icsale ic on ic.finterid=ICe.finterid 
left join (Select * from 
OpenDataSource('Microsoft.ACE.OLEDB.12.0','Data Source="d:\44.xlsx";User ID=admin;Password=;Extended properties=Excel 8.0')...[Sheet1$]
) ss on ss.FBillno=ic.fbillno
where ss.FBillno=ic.fbillno and ice.fentryid=ss.行号

exec sp_configure 'show advanced options',1
reconfigure
exec sp_configure 'Ad Hoc Distributed Queries',1
reconfigure

--开启导入功能
    exec sp_configure 'show advanced options',1
    reconfigure
    exec sp_configure 'Ad Hoc Distributed Queries',1
    reconfigure
    --允许在进程中使用ACE.OLEDB.12
    EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 1
    --允许动态参数
    EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParameters', 1
    
    
    --更新语句1(更新销售发票)
    UPDATE ice set ice.fsourceEntryID=ss.源单分录,ice.fsourcebillNo=ss.源单单号,ice.FSourceInterId=ss.源单内码,ice.FSourceTranType=21
    FROM icsaleentry ice
    INNER JOIN icsale ic ON ic.finterid=ice.finterid
    WHERE /*ic.fbillno=''
    AND*/ ice.fentryid='行号' --@行号{0}
    AND ic.fbillno='发票号码'
    
    
    UPDATE ice set ice.fsourceEntryID=@源单分录{0},ice.fsourcebillNo=@源单单号{0},ice.FSourceInterId=@源单内码{0},ice.FSourceTranType=ss.源单类型--21
    FROM icsaleentry ice
    INNER JOIN icsale ic ON ic.finterid=ice.finterid
    WHERE ice.fentryid='行号' --@行号{0}
    AND ic.fbillno='发票号码'  --@发票号码{0}
    
    --更新语句2(采购发票)
    
update ice set ice.fsourceEntryID=ss.源单分录,ice.fsourcebillNo=ss.源单单号,ice.FSourceInterId=ss.源单内码,ice.FSourceTranType=ss.源单类型--1
from ICPurchaseEntry ice
inner join ICPurchase ic on ic.finterid=ICe.finterid 
left join (Select * from 
OpenDataSource('Microsoft.ACE.OLEDB.12.0','Data Source="d:\22.xlsx";User ID=admin;Password=;Extended properties=Excel 8.0')...[Sheet1$]
) ss on ss.FBillno=ic.fbillno
where ss.FBillno=ic.fbillno and ice.fentryid=ss.行号


UPDATE ice set ice.fsourceEntryID=ss.源单分录,ice.fsourcebillNo=ss.源单单号,ice.FSourceInterId=ss.源单内码,ice.FSourceTranType=ss.源单类型--1
from ICPurchaseEntry ice 
INNER JOIN ICPurchase ic on ic.finterid=ICe.finterid
WHERE ice.fentryid='行号' --@行号{0}
AND ic.fbillno='发票号码'  --@发票号码{0} FBillno



select count(*) as tcount
                                                   from ICPurchaseEntry ice
                                                   inner join ICPurchase ic on ic.finterid=ice.finterid
                                                   where ice.fentryid=1
                                                   and ic.fbillno='ZPOFP025595'


SELECT a.FSourceEntryID AS '源单分录',a.FSourceBillNo AS '源单单号',a.FSourceInterId AS '源单内码',a.FSourceTranType AS '源单类型'
from ICPurchaseEntry a
INNER JOIN ICPurchase b ON a.FInterID=b.FInterID
WHERE a.FEntryID=20
AND b.FBillNo='ZPOFP025596'