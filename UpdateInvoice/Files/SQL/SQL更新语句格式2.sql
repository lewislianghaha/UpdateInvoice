--�ɹ���Ʊ
update ice set ice.fsourceEntryID=ss.Դ����¼,ice.fsourcebillNo=ss.Դ������,ice.FSourceInterId=ss.Դ������,ice.FSourceTranType=1
from ICPurchaseEntry ice
inner join ICPurchase ic on ic.finterid=ICe.finterid 
left join (Select * from 
OpenDataSource('Microsoft.ACE.OLEDB.12.0','Data Source="d:\22.xlsx";User ID=admin;Password=;Extended properties=Excel 8.0')...[Sheet1$]
) ss on ss.FBillno=ic.fbillno
where ss.FBillno=ic.fbillno and ice.fentryid=ss.�к�

update ICPurchase set FStatus=1,FCheckerID=16436 where FDate>='20180501'

--���۷�Ʊ
update ice set ice.fsourceEntryID=ss.Դ����¼,ice.fsourcebillNo=ss.Դ������,ice.FSourceInterId=ss.Դ������,ice.FSourceTranType=21
from icsaleentry ice
inner join icsale ic on ic.finterid=ICe.finterid 
left join (Select * from 
OpenDataSource('Microsoft.ACE.OLEDB.12.0','Data Source="d:\44.xlsx";User ID=admin;Password=;Extended properties=Excel 8.0')...[Sheet1$]
) ss on ss.FBillno=ic.fbillno
where ss.FBillno=ic.fbillno and ice.fentryid=ss.�к�



exec sp_configure 'show advanced options',1
reconfigure
exec sp_configure 'Ad Hoc Distributed Queries',1
reconfigure

--�������빦��
    exec sp_configure 'show advanced options',1
    reconfigure
    exec sp_configure 'Ad Hoc Distributed Queries',1
    reconfigure
    --�����ڽ�����ʹ��ACE.OLEDB.12
    EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 1
    --����̬����
    EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParameters', 1