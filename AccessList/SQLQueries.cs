
namespace AccessList
{
    public static class SQLQueriesTemplates
    {
        public static string SQLEDAccess()
        {
            return @"select 
	Login
    ,Employee
    ,Role
    ,BU
    ,BU_Path
    ,MBU_Name
    ,Location
from (
    SELECT PermissionsAll.LOGIN
          ,Employee.FullNameEngCalculated as Employee
          ,PermissionsAll.role
          ,PermissionsAll.BU
          ,PermissionsAll.BU_path
          ,PermissionsAll.MBU_Name
          ,PermissionsAll.Location
      FROM EnterpriseDirectories.[auth].[PermissionsAll]
      left join EnterpriseDirectories.[app].[Employee] on Employee.login = PermissionsAll.LOGIN and Employee.active = 1
      left join EnterpriseDirectories.[app].[BusinessUnit] as BusinessUnit on BusinessUnit.name = PermissionsAll.BU 
                                                                              #ConditionActiveFBU
      where 
    #Condition
) as PermissionsAll
order by BU_Path
";
        }

        public static string SQLTRMAccess()
        {
            return @"select 
	Login
    ,Employee
    ,Role
    ,BU
    ,BU_Path
    ,ProjectCode
from (
    SELECT PermissionsAll.[LOGIN]
          ,Employee.nameEngfullName as Employee
          ,PermissionsAll.[role]
          ,PermissionsAll.[BU]
          ,PermissionsAll.[BU_path]
          ,PermissionsAll.[ProjectCode]
      FROM TRMSys.[auth].[PermissionsAll]
      left join TRMSys.[app].[Employee] on Employee.login = PermissionsAll.LOGIN and Employee.active = 1
      left join TRMSys.[app].[FinancialBusinessUnit] as BusinessUnit on BusinessUnit.name = PermissionsAll.BU
      where 
    #Condition
) as PermissionsAll
order by BU_Path
";
        }

        public static string SQLInvoicingAccess()
        {
            return @"select 
	Login
    ,Employee
    ,Role
    ,BU
    ,BU_Path
    ,bu_active
    ,LegalEntity
from (
    SELECT PermissionsAll.[LOGIN]
          ,Employee.[fullNameEng] as Employee
          ,PermissionsAll.[role]
          ,PermissionsAll.[BU]
          ,PermissionsAll.[BU_path]
          ,PermissionsAll.[bu_active]
          ,PermissionsAll.[LegalEntity]
      FROM Invoicing.[auth].[PermissionsAll]
      left join Invoicing.[app].[Employee] on Employee.login = PermissionsAll.LOGIN and Employee.active = 1
      left join Invoicing.[app].[BusinessUnit] as BusinessUnit on BusinessUnit.name = PermissionsAll.BU
      where 
    #Condition
) as PermissionsAll
order by BU_Path
";
        }

        public static string SQLCurrentFBUManager()
        {
            return @"use EnterpriseDirectories

select 
	Login	
    ,Employee
    ,Role
    ,BU
    ,BU_Path
from (
    select 
	     Employee.login as Login
        ,Employee.FullNameEngCalculated as Employee
	    --,BusinessUnitEmployeeRole.role as Role
        ,case when BusinessUnitEmployeeRole.role = 1 then 'Manager'
            when BusinessUnitEmployeeRole.role = 2 then 'ManagerDelegated'
            when BusinessUnitEmployeeRole.role = 3 then 'DeliveryManager'
            when BusinessUnitEmployeeRole.role = 4 then 'SnBApprover'
            when BusinessUnitEmployeeRole.role = 5 then 'HorizontalManager'
        else '' end as Role
        ,BusinessUnit.name as BU
        ,BUFullTree.root as BU_Path
    from [app].[BusinessUnitEmployeeRole]
	    join [app].[BusinessUnit] on BusinessUnitEmployeeRole.businessUnitId = businessUnit.Id
        join [app].[BUFullTree] on BusinessUnitEmployeeRole.businessUnitId = BUFullTree.Id
	    join [app].[Employee] on Employee.id = BusinessUnitEmployeeRole.employeeId
    where
        BusinessUnitEmployeeRole.active = 1
        --and BusinessUnitEmployeeRole.role in (1,2)
        and BusinessUnit.active = 1
) as FBUManager
where
    #Condition
order by BU_Path";
        }
        
        public static string SQLTreeCompare()
        {
            return @"use  [EnterpriseDirectories]

--declare @StartDate    date = '20180401',
--		@StartDateOld date = '20180330'
declare @StartDate    date = '#NewTreeDate',
		@StartDateOld date = '#OldTreeDate'
declare	@VersionList  table (VersionStatus int)
insert into @VersionList (VersionStatus) values (6),(7),(8)
;
--version status
--0 'Draft'
--1 'Rejected'
--3 'BOConfirm'
--5 'Approved'	
--6 'Overdue'	
--7 'Applied'
--8 'Failure'
--9 'PreApproved'

-------------------------------------------------
--1. new versions
select
		 BusinessUnitVersion.id
		,BusinessUnitVersion.businessUnitId
		,BusinessUnitVersion.ComingBusinessUnitId
		--,BusinessUnitVersion.versionStatus
        ,case when BusinessUnitVersion.versionStatus = 0 then 'Draft'
		      when BusinessUnitVersion.versionStatus = 1 then 'Rejected'
		      when BusinessUnitVersion.versionStatus = 3 then 'BOConfirm'
		      when BusinessUnitVersion.versionStatus = 5 then 'Approved'	
		      when BusinessUnitVersion.versionStatus = 6 then 'Overdue'	
              when BusinessUnitVersion.versionStatus = 7 then 'Applied'
		      when BusinessUnitVersion.versionStatus = 8 then 'Failure'
		      when BusinessUnitVersion.versionStatus = 9 then'PreApproved'
        else ''
        end as versionStatus
		,BusinessUnitVersion.periodstartDate
		,BusinessUnitVersion.periodendDate
		,BusinessUnitVersion.plannedEndDate
		,BusinessUnitVersion.active
		,BusinessUnitVersion.name
		,BusinessUnitVersion.parentBusinessUnitId
		,BusinessUnitVersion.HorizontalId
		,Horizontal.name as Horizontal
		,BusinessUnitType.name as BusinessUnitType
	into #TempBUVersions
	from (
		select 
			ComingBusinessUnitId
			,max(periodstartDate) as maxStartDate	
		from [app].[BusinessUnitVersion]
		where
			BusinessUnitVersion.versionStatus in (select VersionStatus from @VersionList)
			and (
			(   BusinessUnitVersion.periodstartDate <= @StartDate				
				and BusinessUnitVersion.[plannedEndDate] is null)
			or
			(   BusinessUnitVersion.periodstartDate <= @StartDate				
				and BusinessUnitVersion.[plannedEndDate] > @StartDate))
			
		group by ComingBusinessUnitId ) as BUVersions
	join [app].[BusinessUnitVersion] on BusinessUnitVersion.ComingBusinessUnitId = BUVersions.ComingBusinessUnitId 
									and BusinessUnitVersion.periodstartDate = BUVersions.maxStartDate
									and BusinessUnitVersion.versionStatus in (select VersionStatus from @VersionList)
									and ((   BusinessUnitVersion.periodstartDate <= @StartDate				
											and BusinessUnitVersion.[plannedEndDate] is null)
										or
										(   BusinessUnitVersion.periodstartDate <= @StartDate				
											and BusinessUnitVersion.[plannedEndDate] > @StartDate))
	left join [app].[Horizontal] on BusinessUnitVersion.HorizontalId = Horizontal.Id
	left join [app].[BusinessUnitType] on BusinessUnitVersion.BusinessUnitTypeId = BusinessUnitType.Id

----------------------------------------------------------
--2. old versions
select
		 BusinessUnitVersion.id
		,BusinessUnitVersion.[businessUnitId]
		,BusinessUnitVersion.[ComingBusinessUnitId]
		--,BusinessUnitVersion.[versionStatus]
        ,case when BusinessUnitVersion.versionStatus = 0 then 'Draft'
		      when BusinessUnitVersion.versionStatus = 1 then 'Rejected'
		      when BusinessUnitVersion.versionStatus = 3 then 'BOConfirm'
		      when BusinessUnitVersion.versionStatus = 5 then 'Approved'	
		      when BusinessUnitVersion.versionStatus = 6 then 'Overdue'	
              when BusinessUnitVersion.versionStatus = 7 then 'Applied'
		      when BusinessUnitVersion.versionStatus = 8 then 'Failure'
		      when BusinessUnitVersion.versionStatus = 9 then'PreApproved'
        else ''
        end as versionStatus
		,BusinessUnitVersion.[periodstartDate]
		,BusinessUnitVersion.[periodendDate]     
		,BusinessUnitVersion.[plannedEndDate] 
		,BusinessUnitVersion.[active]
		,BusinessUnitVersion.[name]
		,BusinessUnitVersion.[parentBusinessUnitId]
		,BusinessUnitVersion.HorizontalId
		,Horizontal.name as Horizontal
		,BusinessUnitType.name as BusinessUnitType
	into #TempBUVersionsOld
	from (
		select 
			ComingBusinessUnitId
			,max(periodstartDate) as maxStartDate	
		from [app].[BusinessUnitVersion]
		where
			BusinessUnitVersion.versionStatus in (select VersionStatus from @VersionList)
			and (
			(   BusinessUnitVersion.periodstartDate <= @StartDateOld				
				and BusinessUnitVersion.[plannedEndDate] is null)
			or
			(   BusinessUnitVersion.periodstartDate <= @StartDateOld				
				and BusinessUnitVersion.[plannedEndDate] > @StartDateOld))
			
		group by ComingBusinessUnitId ) as BUVersions
	join [app].[BusinessUnitVersion] on BusinessUnitVersion.ComingBusinessUnitId = BUVersions.ComingBusinessUnitId 
									and BusinessUnitVersion.periodstartDate = BUVersions.maxStartDate
									and BusinessUnitVersion.versionStatus in (select VersionStatus from @VersionList)
									and ((   BusinessUnitVersion.periodstartDate <= @StartDateOld				
											and BusinessUnitVersion.[plannedEndDate] is null)
										or
										(   BusinessUnitVersion.periodstartDate <= @StartDateOld				
											and BusinessUnitVersion.[plannedEndDate] > @StartDateOld))
	left join [app].[Horizontal] on BusinessUnitVersion.HorizontalId = Horizontal.Id
	left join [app].[BusinessUnitType] on BusinessUnitVersion.BusinessUnitTypeId = BusinessUnitType.Id
;

-------------------------------------------------------
--3. new hierarchy
with Hierarchy(id
			,comingBusinessUnitId
			,versionStatus
			,periodstartDate
			,periodendDate
			,name
			,active
			,parentBusinessUnitId
			,businessUnitType
			,horizontalid
			,horizontal
			,Childs
			) as
	(select
		   TempBUVersions.id
		  ,TempBUVersions.comingBusinessUnitId
		  ,TempBUVersions.versionStatus
		  ,TempBUVersions.periodstartDate
		  ,TempBUVersions.periodendDate
		  ,TempBUVersions.name
		  ,TempBUVersions.active
		  ,TempBUVersions.parentBusinessUnitId
		  ,TempBUVersions.businessUnitType
		  ,TempBUVersions.horizontalid
		  ,TempBUVersions.horizontal
		  ,CAST('' AS VARCHAR(MAX))
	  from #TempBUVersions as TempBUVersions
	  where
			parentBusinessUnitId is null			
  
	  UNION ALL

	  select 
		   BUVersions_1.id
		  ,BUVersions_1.ComingBusinessUnitId
		  ,BUVersions_1.versionStatus
		  ,BUVersions_1.periodstartDate
		  ,BUVersions_1.periodendDate
		  ,BUVersions_1.name
		  ,BUVersions_1.active
		  ,BUVersions_1.parentBusinessUnitId
		  ,BUVersions_1.businessUnitType
		  ,BUVersions_1.horizontalid
		  ,BUVersions_1.horizontal
		  ,CAST(CASE WHEN Hierarchy.Childs = '' THEN Hierarchy.Name
		   ELSE Hierarchy.Childs + ' -> ' + Hierarchy.Name
		   END AS VARCHAR(MAX))
	  
	  from #TempBUVersions as BUVersions_1
	  
	  INNER JOIN  Hierarchy as Hierarchy ON BUVersions_1.parentBusinessUnitId = Hierarchy.ComingBusinessUnitId
	)

	select distinct
		 Hierarchy.id		
		,Hierarchy.ComingBusinessUnitId
		,Hierarchy.versionStatus
		,Hierarchy.periodstartDate
		,Hierarchy.periodendDate
		,Hierarchy.name
		,Hierarchy.active
		,Hierarchy.parentBusinessUnitId
		,Hierarchy.businessUnitType
		,Hierarchy.horizontalid
		,Hierarchy.horizontal
		,CASE WHEN Childs = '' THEN name
		  ELSE Childs + ' -> ' + name
		  END as Childs
	into #TreeOnDate
	from Hierarchy
;

-----------------------------------------------
--4. old hierarchy
with Hierarchy(id
			,comingBusinessUnitId
			,versionStatus
			,periodstartDate
			,periodendDate
			,name
			,active
			,parentBusinessUnitId
			,businessUnitType
			,horizontalid
			,horizontal
			,Childs
			) as
	(select
		   TempBUVersions.id
		  ,TempBUVersions.comingBusinessUnitId
		  ,TempBUVersions.versionStatus
		  ,TempBUVersions.periodstartDate
		  ,TempBUVersions.periodendDate
		  ,TempBUVersions.name
		  ,TempBUVersions.active
		  ,TempBUVersions.parentBusinessUnitId
		  ,TempBUVersions.businessUnitType
		  ,TempBUVersions.horizontalid
		  ,TempBUVersions.horizontal
		  ,CAST('' AS VARCHAR(MAX))
	  from #TempBUVersionsOld as TempBUVersions
	  where 			
			parentBusinessUnitId is null
  
	  UNION ALL

	  select 
		   BUVersions_1.id
		  ,BUVersions_1.ComingBusinessUnitId
		  ,BUVersions_1.versionStatus
		  ,BUVersions_1.periodstartDate
		  ,BUVersions_1.periodendDate
		  ,BUVersions_1.name
		  ,BUVersions_1.active
		  ,BUVersions_1.parentBusinessUnitId
		  ,BUVersions_1.businessUnitType
		  ,BUVersions_1.horizontalid
		  ,BUVersions_1.horizontal
		  ,CAST(CASE WHEN Hierarchy.Childs = '' THEN Hierarchy.Name
		   ELSE Hierarchy.Childs + ' -> ' + Hierarchy.Name
		   END AS VARCHAR(MAX))
	  
	  from #TempBUVersionsOld as BUVersions_1

	  INNER JOIN  Hierarchy as Hierarchy ON BUVersions_1.parentBusinessUnitId = Hierarchy.ComingBusinessUnitId
	)

	select distinct
		 Hierarchy.id
		,Hierarchy.ComingBusinessUnitId
		,Hierarchy.versionStatus
		,Hierarchy.periodstartDate
		,Hierarchy.periodendDate
		,Hierarchy.name
		,Hierarchy.active
		,Hierarchy.parentBusinessUnitId
		,Hierarchy.businessUnitType
		,Hierarchy.horizontalid
		,Hierarchy.horizontal
		,CASE WHEN Childs = '' THEN name
		  ELSE Childs + ' -> ' + name
		  END as Childs	
	into #TreeOnDateOld
	from Hierarchy

----------------------------------------------
--5. compare
SELECT 
	 TreeNew.ComingBusinessUnitId as FBUId_New
	,TreeNew.active as FBUActive_New
    ,TreeNew.versionStatus as FBUStatus_New
	,TreeNew.name as FBU_New
	,TreeOld.name as FBU_Old
	,TreeNewParent.name as FBUParent_New
	,TreeOldParent.name as FBUParent_Old
	,TreeNew.businessUnitType as FBUType_New
	,TreeOld.businessUnitType as FBUType_Old
	,TreeNew.Horizontal as Horizontal_New
	,TreeOld.Horizontal as Horizontal_Old
	,TreeNew.Childs as FBU_path_New
	,TreeOld.Childs as FBU_path_old
  from #TreeOnDate as TreeNew
  left join #TreeOnDate as TreeNewParent on TreeNewParent.ComingBusinessUnitId = TreeNew.parentBusinessUnitId

  left join #TreeOnDateOld as TreeOld on TreeOld.ComingBusinessUnitId = TreeNew.ComingBusinessUnitId
  left join #TreeOnDateOld as TreeOldParent on TreeOldParent.ComingBusinessUnitId = TreeOld.parentBusinessUnitId
where
	#Condition
  
order by 
	TreeNew.Childs

--------------------------------------------
--drop temp tables
drop table #TempBUVersions
drop table #TempBUVersionsOld
drop table #TreeOnDate
drop table #TreeOnDateOld

";
        }

        public static string SQLAccessAudit()
        {
            return @"use #DatabaseToUse

declare @PrincipalName varchar(50) = '#Condition'

-------------------------
SELECT  PermissionFilterItemAudit.[id] as PFIId
	   ,PermissionFilterItemAudit.[entityId]
	   ,BusinessUnit.name as FBUName
	   ,EntityType.name as EntityType
	   ,PermissionFilterItemAudit.[permissionId]
       ,Location.name as LocationName   
into #TempPFI
  FROM [authAudit].[PermissionFilterItemAudit]

  left join [auth].[PermissionFilterEntity] as PermissionFilterEntity on PermissionFilterItemAudit.entityId = PermissionFilterEntity.id

  left join [authAudit].[EntityTypeAudit] As EntityType on EntityType.id = PermissionFilterEntity.entityTypeId  
  left join [app].#BusinessUnitTableName on BusinessUnit.id = PermissionFilterEntity.entityId  
  left join [app].Location on Location.id = PermissionFilterEntity.entityId  

  where
	  permissionId in (
						SELECT PermissionAudit.[id]      
						FROM [authAudit].[PermissionAudit]
						join [authAudit].[PrincipalAudit] on PrincipalAudit.id = PermissionAudit.principalId  
						 where
						   PrincipalAudit.name = @PrincipalName
					   )
   
-------------------------   
SELECT PermissionFilterItemDeleted.[id] as PFIId
	  ,AuditRevisionEntityPFI.Author as DeletedPFIBy
	  ,AuditRevisionEntityPFI.RevisionDate as DeletePFIDate
	  ,TempPFI.FBUName
      ,TempPFI.LocationName
	  ,PermissionFilterItemDeleted.[permissionId]
into #TempDeletedPFI
  FROM [authAudit].[PermissionFilterItemAudit] as PermissionFilterItemDeleted
  left join [authAudit].[AuditRevisionEntity] as AuditRevisionEntityPFI on PermissionFilterItemDeleted.Rev = AuditRevisionEntityPFI.id
  left join #TempPFI as TempPFI on TempPFI.PFIId = PermissionFilterItemDeleted.id
where 
	PermissionFilterItemDeleted.id in (select PFIId from #TempPFI)
	and PermissionFilterItemDeleted.REVTYPE = 2

-------------------------
SELECT distinct
	 Principal.name	as Principal
	,BusinessRole.name as Role
	,isnull(Permission.comment, '') as comment
	,AuditRevisionEntity.Author as [Deleted permission by]
	,AuditRevisionEntity.RevisionDate as [Delete permission date]
	,TempPFI.PermissionId
	,TempPFI.EntityType
	,TempPFI.FBUName
    ,TempPFI.LocationName
	,TempDeletedPFI.DeletedPFIBy as [Deleted PFI by]
	,TempDeletedPFI.DeletePFIDate as [Delete PFI date]
    ,Permission.[createdBy]
    ,Permission.[createDate]
    ,Permission.[modifiedBy]
    ,Permission.[modifyDate]

  FROM [authAudit].[PermissionAudit] as Permission
  join [authAudit].[PrincipalAudit] as Principal on Principal.id = Permission.principalId  
  join [auth].[BusinessRole] as BusinessRole on BusinessRole.id = Permission.roleId  

  left join [authAudit].[PermissionAudit] as PermissionDeleted on Permission.Id = PermissionDeleted.id and PermissionDeleted.RevType = 2
  left join [authAudit].[AuditRevisionEntity] as AuditRevisionEntity on PermissionDeleted.Rev = AuditRevisionEntity.id
  
  left join [authAudit].[PermissionFilterItemAudit] as PermissionFilterItem on PermissionFilterItem.permissionId = Permission.id

  left join #TempPFI as TempPFI on TempPFI.PFIId = PermissionFilterItem.id
  left join #TempDeletedPFI as TempDeletedPFI on TempDeletedPFI.PFIId = TempPFI.PFIId  

where
	Principal.name = @PrincipalName

order by
	Principal.name
	,BusinessRole.name
    ,TempPFI.PermissionId
    ,TempPFI.EntityType
	,TempDeletedPFI.DeletePFIDate


drop table #TempPFI
drop table #TempDeletedPFI
";
        }

        public static string SQLFBUManagerAudit()
        {
            return @"use [EnterpriseDirectories]

select
      BusinessUnitEmployeeRoleAudit.id      
into #TempRolesId
  FROM [appAudit].[BusinessUnitEmployeeRoleAudit]
  join [app].[Employee] on BusinessUnitEmployeeRoleAudit.employeeId = Employee.id
  join [app].[BusinessUnit] on BusinessUnitEmployeeRoleAudit.businessUnitId = businessUnit.Id
  where
    #Condition
	--Employee.login = @PrincipalName
	--businessUnit.name = 'Abu Dhabi Investment Authority'

SELECT 
	  OldFBU.FBU
	  ,OldFBU.active as FBUActive	  
	  ,OldFBU.login
	  ,OldFBU.Employee
	  ,case when OldFBU.role = 1 then 'Manager'
            when OldFBU.role = 2 then 'ManagerDelegated'
            when OldFBU.role = 5 then 'HorizontalManager'
        else cast(OldFBU.role as varchar(10)) end as Role
      --,BusinessUnit.active as FBUActive
	  --,Employee.login as Login
	  --,Employee.fullnameengcalculated as Employee
	  --,BusinessUnitEmployeeRoleAudit.[role]
	  ,case when BusinessUnitEmployeeRoleAudit.[REVTYPE] = 2 then AuditRevisionEntity.[RevisionDate] else null end as DeleteDate
	  ,case when BusinessUnitEmployeeRoleAudit.[REVTYPE] = 2 then AuditRevisionEntity.[Author] else null end as DeletedBy
      ,BusinessUnitEmployeeRoleAudit.[createdBy]      
      ,BusinessUnitEmployeeRoleAudit.[createDate]      
      --,BusinessUnitEmployeeRoleAudit.[modifyDate]      
      --,BusinessUnitEmployeeRoleAudit.[modifiedBy]
  FROM [appAudit].[BusinessUnitEmployeeRoleAudit]
  join [appAudit].[AuditRevisionEntity] on AuditRevisionEntity.id = BusinessUnitEmployeeRoleAudit.Rev
  left join [app].[Employee] on BusinessUnitEmployeeRoleAudit.employeeId = Employee.id
  left join [app].[BusinessUnit] on BusinessUnitEmployeeRoleAudit.businessUnitId = businessUnit.Id
  left join 
  (select distinct
	OldRoles.id
	,OldRoles.role
	,OldFBU.name as FBU
	,oldEmployee.login as Login
	,oldEmployee.fullnameengcalculated as Employee	  	
	,OldFBU.active
  from [appAudit].[BusinessUnitEmployeeRoleAudit] as OldRoles
  left join [app].[BusinessUnit] as OldFBU on OldRoles.businessUnitId = OldFBU.Id
  left join [app].[Employee] as oldEmployee on OldRoles.employeeId = oldEmployee.id
  where
	OldRoles.id in (SELECT id from #TempRolesId)
	and OldRoles.[REVTYPE] != 2
	and OldFBU.id is not null
  ) as OldFBU on OldFBU.id = BusinessUnitEmployeeRoleAudit.id
  where 
	BusinessUnitEmployeeRoleAudit.id in (SELECT id from #TempRolesId)
order by
	OldFBU.FBU
	,OldFBU.login
	,case when BusinessUnitEmployeeRoleAudit.[REVTYPE] = 2 then AuditRevisionEntity.[RevisionDate] else null end

drop table #TempRolesId";
        }

        public static string SQLFBUVersionSearch()
        {
            return @"use [EnterpriseDirectories]

SELECT
	  [comingBusinessUnitId] as FBUId
      ,[active] as FBUActive
      ,[name] FBUName
	  ,cast([periodstartDate] as date) as StartDate
	  ,cast([periodendDate] as date) as EndDate
      ,case when BusinessUnitVersion.versionStatus = 0 then 'Draft'
            when BusinessUnitVersion.versionStatus = 1 then 'Rejected'
            when BusinessUnitVersion.versionStatus = 3 then 'BOConfirm'
            when BusinessUnitVersion.versionStatus = 5 then 'Approved'	
            when BusinessUnitVersion.versionStatus = 6 then 'Overdue'	
            when BusinessUnitVersion.versionStatus = 7 then 'Applied'
            when BusinessUnitVersion.versionStatus = 8 then 'Failure'
            when BusinessUnitVersion.versionStatus = 9 then 'PreApproved'
            else '' end as VersionStatus
      ,businessUnitChangeBatchId as BatchID
  FROM [app].[BusinessUnitVersion]
  where
  businessUnitId in (SELECT [comingBusinessUnitId]
                    FROM [app].[BusinessUnitVersion]  
                    where  name like '#Condition%')";
        }
    }
}
