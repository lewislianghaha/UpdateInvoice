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
    
    
    --�������1(�������۷�Ʊ)
    UPDATE ice set ice.fsourceEntryID=ss.Դ����¼,ice.fsourcebillNo=ss.Դ������,ice.FSourceInterId=ss.Դ������,ice.FSourceTranType=21
    FROM icsaleentry ice
    INNER JOIN icsale ic ON ic.finterid=ice.finterid
    WHERE /*ic.fbillno=''
    AND*/ ice.fentryid='�к�' --@�к�{0}
    AND ic.fbillno='��Ʊ����'
    
    
    UPDATE ice set ice.fsourceEntryID=@Դ����¼{0},ice.fsourcebillNo=@Դ������{0},ice.FSourceInterId=@Դ������{0},ice.FSourceTranType=ss.Դ������--21
    FROM icsaleentry ice
    INNER JOIN icsale ic ON ic.finterid=ice.finterid
    WHERE ice.fentryid='�к�' --@�к�{0}
    AND ic.fbillno='��Ʊ����'  --@��Ʊ����{0}
    
    --�������2(�ɹ���Ʊ)
    
update ice set ice.fsourceEntryID=ss.Դ����¼,ice.fsourcebillNo=ss.Դ������,ice.FSourceInterId=ss.Դ������,ice.FSourceTranType=ss.Դ������--1
from ICPurchaseEntry ice
inner join ICPurchase ic on ic.finterid=ICe.finterid 
left join (Select * from 
OpenDataSource('Microsoft.ACE.OLEDB.12.0','Data Source="d:\22.xlsx";User ID=admin;Password=;Extended properties=Excel 8.0')...[Sheet1$]
) ss on ss.FBillno=ic.fbillno
where ss.FBillno=ic.fbillno and ice.fentryid=ss.�к�


UPDATE ice set ice.fsourceEntryID=ss.Դ����¼,ice.fsourcebillNo=ss.Դ������,ice.FSourceInterId=ss.Դ������,ice.FSourceTranType=ss.Դ������--1
from ICPurchaseEntry ice 
INNER JOIN ICPurchase ic on ic.finterid=ICe.finterid
WHERE ice.fentryid='�к�' --@�к�{0}
AND ic.fbillno='��Ʊ����'  --@��Ʊ����{0} FBillno



select count(*) as tcount
                                                   from ICPurchaseEntry ice
                                                   inner join ICPurchase ic on ic.finterid=ice.finterid
                                                   where ice.fentryid=1
                                                   and ic.fbillno='ZPOFP025595'


SELECT a.FSourceEntryID AS 'Դ����¼',a.FSourceBillNo AS 'Դ������',a.FSourceInterId AS 'Դ������',a.FSourceTranType AS 'Դ������'
from ICPurchaseEntry a
INNER JOIN ICPurchase b ON a.FInterID=b.FInterID
WHERE a.FEntryID=20
AND b.FBillNo='ZPOFP025596'