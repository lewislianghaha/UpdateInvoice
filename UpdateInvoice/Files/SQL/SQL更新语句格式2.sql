--采购发票
update ice set ice.fsourceEntryID=ss.源单分录,ice.fsourcebillNo=ss.源单单号,ice.FSourceInterId=ss.源单内码,ice.FSourceTranType=1
from ICPurchaseEntry ice
inner join ICPurchase ic on ic.finterid=ICe.finterid 
left join (Select * from 
OpenDataSource('Microsoft.ACE.OLEDB.12.0','Data Source="d:\22.xlsx";User ID=admin;Password=;Extended properties=Excel 8.0')...[Sheet1$]
) ss on ss.FBillno=ic.fbillno
where ss.FBillno=ic.fbillno and ice.fentryid=ss.行号

update ICPurchase set FStatus=1,FCheckerID=16436 where FDate>='20180501'

--销售发票
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