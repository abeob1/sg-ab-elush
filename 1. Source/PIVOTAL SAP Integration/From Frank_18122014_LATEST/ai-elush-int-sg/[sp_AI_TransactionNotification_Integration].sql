USE [SAP_Elush]
GO
/****** Object:  StoredProcedure [dbo].[sp_AI_TransactionNotification_Integration]    Script Date: 06/20/2013 10:28:12 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROC [dbo].[sp_AI_TransactionNotification_Integration]
    @object_type NVARCHAR(20), 				-- SBO Object Type
    @transaction_type NCHAR(1),				-- [A]dd, [U]pdate, [D]elete, [C]ancel, C[L]ose
    @num_of_cols_in_key INT,
    @list_of_key_cols_tab_del NVARCHAR(255),
    @list_of_cols_val_tab_del NVARCHAR(255),
    @error INT OUTPUT,
    @error_message NVARCHAR(200) OUTPUT
AS 
BEGIN
	declare @IntegrationSAPUser nvarchar(8)
	set @IntegrationSAPUser='manager'
	declare @ID numeric(18,0) --auto ID return
	
	---------------------------Purchase Order-------------
	if @object_type='22'
	begin
		---------------------------Not allow multiple warehouse on PO-------------
		if (select count(WhsCode) from (select distinct WhsCode from POR1  with(nolock) where docentry=@list_of_cols_val_tab_del) T0) >1
		begin
			select @error=-1, @error_message='Multiple Warehouse in PO is not allowed!!!'
			return
		end
		
		---------------------------Send to Integration-------------
		if exists(select 1 from POR1 T0  with(nolock)  
			join OWHS T1  with(nolock) on T0.WhsCode=T1.WhsCode
			where T0.DocEntry=@list_of_cols_val_tab_del and isnull(T1.U_POSFlag,'N')='Y') --only for POS
		begin
			---------------PO header-------------
			insert into SAP_Integration..POHeader --Hardcode, fix database name: SAP_Integration
			select T0.DocEntry,T0.CardCode,T0.CardName,T0.DocDate,T0.DocDueDate,T0.CANCELED,
				CASE when T0.DocStatus='C' then 'Y' else 'N' end, --isclose
				GETDATE(),null, null,N'SAP',NumAtCard
			from OPOR T0 with(nolock)  where T0.DocEntry=@list_of_cols_val_tab_del
			---------------Get ID-----------------
			set @ID=@@IDENTITY
			---------------PO Line-------------
			insert into SAP_Integration..POLine --Hardcode, fix database name: SAP_Integration			
			select @ID,T0.LineNum,T0.ItemCode,T0.Dscription,T0.Quantity,T0.WhsCode,T1.ManSerNum
			from POR1 T0  with(nolock) 
			join OITM T1 with(nolock) on T0.ItemCode=T1.ItemCode
			where T0.DocEntry=@list_of_cols_val_tab_del
		end	
		
		--------------SEND EMAIL-----------------
		if exists(select 1 from OPOR where docentry=@list_of_cols_val_tab_del and ISNULL(U_Type,'')='NP')
			and @transaction_type<>'C'
		begin
			---------------PO header-------------
			insert into SAP_Integration..SendEmailPOHeader
			select T0.DocEntry,T0.DocDate, T0.CardCode,T0.CardName,T0.VatSum,T0.DocTotal-T0.VatSum,T0.DocTotal,GETDATE(),null,null 
			from OPOR T0
			where T0.DocEntry=@list_of_cols_val_tab_del
			---------------Get ID-----------------
			set @ID=@@IDENTITY
			---------------PO Line-------------
			insert into SAP_Integration..SendEmailPOLine
			select @ID,T1.ItemCode,T1.Dscription,T1.Quantity,T1.Price,T1.LineTotal, T1.WhsCode
			from  POR1 T1 
			where T1.DocEntry=@list_of_cols_val_tab_del
		end
		
		--------------Update Total Quantity-----------------
		declare @TotalQty numeric(19,6)
		declare @TotalOpenQty numeric(19,6)
		
		select @TotalQty=SUM(Quantity),@TotalOpenQty=SUM(OpenCreQty) 
		from POR1 where DocEntry=@list_of_cols_val_tab_del
		
		update OPOR set U_TotalQty=@TotalQty,U_TotalOpenQty=@TotalOpenQty
		where DocEntry=@list_of_cols_val_tab_del
	end
	
	---------------------------Goods Receipt PO-------------
	if @object_type='20'  and @transaction_type='A'
	begin
		--------------Update Total Quantity-----------------
		declare @TotalQty1 numeric(19,6)
		declare @TotalOpenQty1 numeric(19,6)
		
		select @TotalQty1=SUM(Quantity),@TotalOpenQty1=SUM(OpenCreQty) 
		from PDN1 where DocEntry=@list_of_cols_val_tab_del
		
		update OPDN set U_TotalQty=@TotalQty1,U_TotalOpenQty=@TotalOpenQty1 
		where DocEntry=@list_of_cols_val_tab_del
		
		--------------Update Total Open Quantity of PO-----------------
		update T0 set U_TotalOpenQty=(select SUM(OpenCreQty) from POR1 T1 where T1.DocEntry=T0.DocEntry)
		from OPOR T0		
		where T0.DocEntry in 
		(
			select T1.DocEntry from PDN1 T0
			join POR1 T1 on T0.BaseEntry=T1.DocEntry and T0.BaseLine=T1.LineNum and T0.BaseType='22'
			where T0.DocEntry=@list_of_cols_val_tab_del 
		)
	end
	
	---------------------------AP Invoice-------------
	if @object_type='18'
	begin
		--------------Update Total Open Quantity of GRPO-----------------
		update T0 set U_TotalOpenQty=(select SUM(OpenCreQty) from PDN1 T1 where T1.DocEntry=T0.DocEntry)
		from OPDN T0		
		where T0.DocEntry in 
		(
			select T1.DocEntry from PCH1 T0
			join PDN1 T1 on T0.BaseEntry=T1.DocEntry and T0.BaseLine=T1.LineNum and T0.BaseType='20'
			where T0.DocEntry=@list_of_cols_val_tab_del 
		)
	end
	
	---------------------------Item Master Data-------------------------------
	if @object_type='4' 
	begin
		insert into SAP_Integration..ItemMasterData  --Hardcode, fix database name: SAP_Integration
		select T0.ItemCode, T0.ItemName, T0.ItmsGrpCod, T2.ItmsGrpNam,
			T0.ItemCode,T0.ManSerNum,isnull(T0.CardCode,'Elush'),T0.SuppCatNum,T0.InvntItem, T1.Price,
			T0.U_Group,T0.U_Family,T0.U_Model,T0.U_ScreenSize,T0.U_HDCapacity,T0.U_Processor,
			T0.U_RAMSize,T0.U_Generation,T0.U_Network,T0.U_Kind,T0.U_ForModel,
			T0.U_Watt,T0.U_Type,T0.U_Specification,T0.U_Fits,T0.U_Color,T0.U_Brand,T0.U_AltCode1,
			T0.U_AltCode2,T0.U_AltCode3,T0.U_MarginRank,T0.U_Status,
			CASE when @transaction_type='D' then 'Y' else 'N' end, frozenfor,--inactive
			GETDATE(),null, null,N'SAP'
		from OITM T0  with(nolock) 
		join ITM1 T1  with(nolock) on T0.ItemCode=T1.ItemCode and T1.PriceList=1 --hardcode, fix pricelist 1
		join OITB T2  with(nolock) on T2.ItmsGrpCod=T0.ItmsGrpCod
		where T0.ItemCode=@list_of_cols_val_tab_del and isnull(T0.U_POSFlag,'N')='Y' --only for POS Item
		
		if @transaction_type='D'
		begin
			insert into SAP_Integration..ItemMasterData  --Hardcode, fix database name: SAP_Integration
			select @list_of_cols_val_tab_del,'', 0, '',
			'','Y','','','Y', 0,'','','','','','','','','','','',
			'','','','','','','','','','','','Y','N',GETDATE(),null, null,N'SAP'
		end
	end
	---------------------------Business Partner-------------------------------
	if @object_type='2'
	begin
		insert into SAP_Integration..BusinessParterMaster --Hardcode, fix database name: SAP_Integration
		select top(1) -- incase bill-to or ship-to return more than 2 records
			T0.CardCode,T0.CardName, T0.GroupCode,T1.GroupName,
			T2.Address,T2.Block,T2.City,T2.ZipCode,T2.Country,
			T3.Address,T3.Block,T3.City,T3.ZipCode,T3.Country,
			T0.Phone1,T0.Fax,T0.E_Mail,T0.CntctPrsn,
			CASE when @transaction_type='D' then 'Y' else 'N' end, frozenfor,--inactive
			GETDATE(),null, null,N'SAP',
			T0.CardType,T2.Street,T3.Street
		from OCRD T0  with(nolock) 
		join OCRG T1 with(nolock)  on T0.GroupCode=T1.GroupCode
		left join CRD1 T2  with(nolock) on T2.CardCode=T0.CardCode and T2.AdresType='B' -- Bill-To
		left join CRD1 T3  with(nolock) on T3.CardCode=T0.CardCode and T3.AdresType='S' -- Ship-To
		where T0.CardCode=@list_of_cols_val_tab_del and isnull(T0.U_POSFlag,'N')='Y' --only for POS BP
		
		if @transaction_type='D'
		begin
			insert into SAP_Integration..BusinessParterMaster --Hardcode, fix database name: SAP_Integration
			select @list_of_cols_val_tab_del,'', 0,'','','','','','',
			'','','','','',	'','','','', 'Y', 'N',	GETDATE(),null, null,N'SAP','','',''
		end
	end
	---------------------------Inventory Transfer-------------------------------
	if @object_type='67' and @transaction_type='A'
	begin
		if exists(select 1 from OWTR T0  with(nolock)  
			join OWHS T1  with(nolock) on isnull(T0.U_ToStore,'')=T1.WhsCode
			join OUSR T2 on T2.USERID=T0.UserSign
			where T0.DocEntry=@list_of_cols_val_tab_del and isnull(T1.U_POSFlag,'N')='Y' --only for POS
			and T0.Filler = 'HO' and  T2.USER_CODE<>@IntegrationSAPUser ) --avoid duplicate transaction from POS, check by user
		begin
			---------------Transfer header-------------
			insert into SAP_Integration..TransferHeader --Hardcode, fix database name: SAP_Integration
			select T0.DocEntry,'',T0.Filler,T0.DocDate,'SAP',T0.U_ToStore, 
				GETDATE(),null, null
			from OWTR T0 with(nolock)  where T0.DocEntry=@list_of_cols_val_tab_del
			---------------Get ID-----------------
			set @ID=@@IDENTITY
			---------------Transfer Line-------------
			insert into SAP_Integration..TransferLine --Hardcode, fix database name: SAP_Integration			
			select @ID,T0.ItemCode,T0.Dscription,T0.WhsCode,T0.Quantity,T0.LineNum,T1.ManSerNum
			from WTR1 T0  with(nolock)
			join OITM T1 with(nolock) on T0.ItemCode=T1.ItemCode
			where T0.DocEntry=@list_of_cols_val_tab_del
			---------------Transfer Serial Number-------------
			insert into SAP_Integration..SerialNumber --Hardcode, fix database name: SAP_Integration			
			select distinct @object_type,T2.ItemCode,T2.DistNumber,'',@list_of_cols_val_tab_del,T0.DocLine,@ID,''
			from OITL T0 with(nolock) 															--inventory transaction log
			join ITL1 T1  with(nolock) on T0.LogEntry=T1.LogEntry									--serial detail in transaction
			join OSRN T2  with(nolock) on T2.SysNumber=T1.SysNumber and T2.ItemCode=T1.ItemCode	--serial master, get serial number
			where T0.DocEntry=@list_of_cols_val_tab_del and T0.DocType=@object_type --and T0.DocLine=0
		end	
	end

	---------------------------Goods Receipt-------------------------------
	if @object_type='59'  and @transaction_type='A'
	begin
		if exists(select 1 from IGN1 T0  with(nolock)  
			join OWHS T1  with(nolock) on T0.WhsCode=T1.WhsCode
			where T0.DocEntry=@list_of_cols_val_tab_del and isnull(T1.U_POSFlag,'N')='Y') --only for POS
		begin
			---------------Goods Receipt header-------------
			insert into SAP_Integration..GoodsReceiptHeader --Hardcode, fix database name: SAP_Integration
			select T0.DocEntry,T0.DocDate,T0.TaxDate,
				GETDATE(),null, null,N'SAP'
			from OIGN T0 with(nolock)  where T0.DocEntry=@list_of_cols_val_tab_del
			---------------Goods Receipt Line-------------
			insert into SAP_Integration..GoodsReceiptLine --Hardcode, fix database name: SAP_Integration			
			select T0.DocEntry,T0.ItemCode,T0.Dscription,T0.WhsCode,T0.Quantity,T1.ManSerNum, T0.LineNum
			from IGN1 T0  with(nolock) 
			join OITM T1  with(nolock) on T0.ItemCode=T1.ItemCode
			where T0.DocEntry=@list_of_cols_val_tab_del
			---------------Serial Number ------------------------
			insert into SAP_Integration..SerialNumber --Hardcode, fix database name: SAP_Integration			
			select distinct @object_type,T2.ItemCode,T2.DistNumber,'',@list_of_cols_val_tab_del,T0.DocLine,@list_of_cols_val_tab_del,''
			from OITL T0 with(nolock) 															--inventory transaction log
			join ITL1 T1  with(nolock) on T0.LogEntry=T1.LogEntry									--serial detail in transaction
			join OSRN T2  with(nolock) on T2.SysNumber=T1.SysNumber and T2.ItemCode=T1.ItemCode	--serial master, get serial number
			where T0.DocEntry=@list_of_cols_val_tab_del and T0.DocType=@object_type --and T0.DocLine=0
		end	
	end
	---------------------------Goods Issue-------------------------------
	if @object_type='60'  and @transaction_type='A'
	begin
		if exists(select 1 from IGE1 T0  with(nolock)  
			join OWHS T1  with(nolock) on T0.WhsCode=T1.WhsCode
			where T0.DocEntry=@list_of_cols_val_tab_del and isnull(T1.U_POSFlag,'N')='Y') --only for POS
		begin
			---------------Goods Issue header-------------
			insert into SAP_Integration..GoodsIssueHeader --Hardcode, fix database name: SAP_Integration
			select T0.DocEntry,T0.DocDate,T0.TaxDate,
				GETDATE(),null, null,N'SAP'
			from OIGE T0 with(nolock)  where T0.DocEntry=@list_of_cols_val_tab_del
			---------------Goods Issue Line-------------
			insert into SAP_Integration..GoodsIssueLine --Hardcode, fix database name: SAP_Integration			
			select T0.DocEntry,T0.ItemCode,T0.Dscription,T0.WhsCode,T0.Quantity,T1.ManSerNum, T0.linenum	
			from IGE1 T0  with(nolock) 
			join OITM T1 with(nolock) on T0.Itemcode=T1.ItemCode
			where T0.DocEntry=@list_of_cols_val_tab_del
			---------------Serial Number ------------------------
			insert into SAP_Integration..SerialNumber --Hardcode, fix database name: SAP_Integration			
			select distinct @object_type,T2.ItemCode,T2.DistNumber,'',@list_of_cols_val_tab_del,T0.DocLine,@list_of_cols_val_tab_del,''
			from OITL T0 with(nolock) 															--inventory transaction log
			join ITL1 T1  with(nolock) on T0.LogEntry=T1.LogEntry									--serial detail in transaction
			join OSRN T2  with(nolock) on T2.SysNumber=T1.SysNumber and T2.ItemCode=T1.ItemCode	--serial master, get serial number
			where T0.DocEntry=@list_of_cols_val_tab_del and T0.DocType=@object_type --and T0.DocLine=0
		end	
	end
	---------------------------Stock take-------------------------------
	if @object_type='10000071' and @transaction_type='A'
	begin
		if exists(select 1 from IQR1 T0  with(nolock)  
			join OWHS T1  with(nolock) on T0.WhsCode=T1.WhsCode
			where T0.DocEntry=@list_of_cols_val_tab_del and isnull(T1.U_POSFlag,'N')='Y') --only for POS
		begin
			
			--------------Stock take-------------
			insert into SAP_Integration..StockTake --Hardcode, fix database name: SAP_Integration
			select T0.DocEntry,T1.DocLineNum,T1.ItemCode,T1.ItemName,T1.Quantity,T1.WhsCode,T0.DocDate,
				GETDATE(),null, null,N'SAP',T2.ManSerNum
			from OIQR T0 with(nolock) 
			join IQR1 T1  with(nolock) on T0.DocEntry=T1.DocEntry
			join OITM T2 with(nolock) on T2.ItemCode=T1.ItemCode
			where T0.DocEntry=@list_of_cols_val_tab_del  and T1.Quantity<>0
			---------------Get ID-----------------
			set @ID=@@IDENTITY
			--select @error=-1, @error_message='here'
			--return
			---------------Serial Number ------------------------
			insert into SAP_Integration..SerialNumber --Hardcode, fix database name: SAP_Integration			
			select distinct @object_type,T2.ItemCode,T2.DistNumber,'',@list_of_cols_val_tab_del,T0.DocLine,@list_of_cols_val_tab_del,''
			from OITL T0 with(nolock) 															--inventory transaction log
			join ITL1 T1  with(nolock) on T0.LogEntry=T1.LogEntry									--serial detail in transaction
			join OSRN T2  with(nolock) on T2.SysNumber=T1.SysNumber and T2.ItemCode=T1.ItemCode	--serial master, get serial number
			where T0.DocEntry=@list_of_cols_val_tab_del and T0.DocType=@object_type --and T0.DocLine=0
		end
	end
	
	
	select @error, @error_message
END                                               
                                              
