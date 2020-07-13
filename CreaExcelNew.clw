                    Member()
 !Clase creada por victor montanes
 !Esta clase se provee tal y como es sin ninguna garantia
 
                    Map
                    End

    INCLUDE('SystemString.inc'),ONCE
    INCLUDE('CreaExcelNew.inc'),ONCE
    INCLUDE('xlsxwriter.inc'),ONCE


OwnerSql            STRING(200)
!Definimos una tabla boba para hacer el query contra la DB
SQLFile             FILE,DRIVER('MSSQL','/TURBOSQL=TRUE'),OWNER(OwnerSql),NAME('dbo.SQuery'),PRE(SQ1),BINDABLE,THREAD !                     
Record                  RECORD,PRE()
C1                          CSTRING(200)                                       
C2                          CSTRING(200)                                       
C3                          CSTRING(200)                                       
C4                          CSTRING(200)                                       
C5                          CSTRING(200)                                       
C6                          CSTRING(200)                                       
C7                          CSTRING(200)                                       
C8                          CSTRING(200)                                       
C9                          CSTRING(200)                                       
C10                         CSTRING(200)                                       
C11                         CSTRING(200)                                       
C12                         CSTRING(200)                                       
C13                         CSTRING(200)                                       
C14                         CSTRING(200)                                       
C15                         CSTRING(200)                                       
C16                         CSTRING(200)                                       
C17                         CSTRING(200)                                       
C18                         CSTRING(200)                                       
C19                         CSTRING(200)                                       
C20                         CSTRING(200)                                       
C21                         CSTRING(200)         
C22                         CSTRING(200)         
C23                         CSTRING(200)         
C24                         CSTRING(200)         
C25                         CSTRING(200)         
C26                         CSTRING(200)         
C27                         CSTRING(200)         
C28                         CSTRING(200)         
C29                         CSTRING(200)         
C30                         CSTRING(200) 
C31                         CSTRING(200)         
C32                         CSTRING(200)         
C33                         CSTRING(200)         
C34                         CSTRING(200)         
C35                         CSTRING(200)         
C36                         CSTRING(200)         
C37                         CSTRING(200)         
C38                         CSTRING(200)         
                        END
                    END    
    

    
!----------------------------------------------------------
!
!----------------------------------------------------------
CreaExcelClassNew.Construct PROCEDURE
    CODE
        Self.QCampos   &= NEW CamposQueue
        SELF.QExport   &= NEW QExportBase
!----------------------------------------------------------
!
!----------------------------------------------------------
CreaExcelClassNew.Destruct  PROCEDURE
    CODE
        FREE(SELF.QCampos)
        FREE(SELF.QExport)
        DISPOSE(SELF.QCampos)
        DISPOSE(SELF.QExport)
!---------------------------------------------------------
!    
!---------------------------------------------------------
CreaExcelClassNew.Init      PROCEDURE(STRING pFnOutput,STRING pOwner,byte Abrirarchi=1)
    CODE
        SELF.FnOutPut    = pFnOutPut
        SELF.FnOwnerName = pOwner
        self.Abrirarchi =Abrirarchi
		
          
!---------------------------------------------------------
!    Generar Reporte a excel
!---------------------------------------------------------        
CreaExcelClassNew.GenerarReporte    PROCEDURE(STRING pQuery,STRING pTitulo,STRING pHeaders, BYTE pTotaliza,<string pAgrupa>)
CAMPO                                   CSTRING(200)
oStr                                    SystemStringClass
oStr2                                   SystemStringClass
ObjXls                                  &xlsxwriter
Primero                                 BYTE
Mismo                                   CSTRING(40)


pgWindow                                WINDOW,AT(,,306,28),FONT('Segoe UI',10),CENTER,COLOR(00C0FFFFh),GRAY
                                            STRING('Exportando a Excel por favor espere'),AT(2,2,303),USE(?STRING999),FONT(,,00CC6600h, |
                                                FONT:regular),CENTER
                                            PROGRESS,AT(7,14,297),USE(?PROGRESS999),RANGE(0,100)
                                        END


    CODE
        OPEN(pgWindow)
        DISPLAY
        if clip(pAgrupa)<>'' then self.NombreAgrupador=pAgrupa END
        SELF.SQuery = UPPER(pQuery)
        
        !Descomponempos el query para obtener los campos del mismo
        self.LLenaCampos(pHeaders)
        OwnerSql =CLIP(self.FnOwnerName)
        OPEN(SQLFile)
        
        
        
        SQLFile{PROP:SQL}=SELF.SQCount
        
        LOOP 
            NEXT(SQLFile)
           
            IF ERRORCODE() THEN BREAK END
            ?PROGRESS999{PROP:RangeHigh}=SQ1:C1
        END    
        ObjXls&=NEW(xlsxwriter) 
        ObjXls.NewWorkbook(SELF.FnOutPut)
        worksheet#=ObjXls.AddSheet('Sheet1')
        !-----------Parametros de la hoja-------------------
        col#=1;ren#=1
        ObjXls.ClearFormat()
        ObjXls.SetSelection(1,1,2,2)
        ObjXls.Format.FontSize=14
        ObjXls.Format.FontStyle=FONT:Bold
        ObjXls.SetFormat()
        oStr2.Str(clip(pTitulo))
        ostr2.Split(',')
        LOOP i#=1 to  oStr2.GetLinesCount()
            x#=  ObjXls.Merge(ren#,1,ren#,SELF.ContCampos)
            err#=ObjXls.WriteString(ren#,col#,clip(oStr2.GetLineValue(i#)));ren#+=1
        END
        ObjXls.ClearFormat()
        ObjXls.SetSelection(1,1,2,2)
        ObjXls.Format.FontSize=12
        ObjXls.Format.FontStyle=FONT:Bold
        ObjXls.Format.Color=00F2E4D7h
        ObjXls.SetFormat()
        ren#=3
        LOOP i#=1 TO SELF.ContCampos
            GET(SELF.QCampos,i#)
            if ~ERRORCODE()
                err#=ObjXls.WriteString(ren#,col#,CLIP(SELF.QCampos.Nombre));col#+=1
            END        
        END
        ren#=4
        !BUFFER(queryview,200) 
        CLEAR(SQ1:Record)
        SQLFile{PROP:SQL}=SELF.SQuery
        LOOP 
            !NEXT(SQLFile)
            NEXT(SQLFile)
            IF ERRORCODE() THEN BREAK END
            if self.PosicionAgrupador
                COL# = 1
                CAMPO = CHOOSE(self.PosicionAgrupador,SQ1:C1,SQ1:C2,SQ1:C3,SQ1:C4,SQ1:C5,SQ1:C6,SQ1:C7,SQ1:C8,SQ1:C9,SQ1:C10,SQ1:C11,SQ1:C12,SQ1:C13,SQ1:C14,SQ1:C15,SQ1:C16,SQ1:C17,SQ1:C18,SQ1:C19,SQ1:C20,SQ1:C21,SQ1:C22,SQ1:C23,SQ1:C24,SQ1:C25,SQ1:C26,SQ1:C27,SQ1:C28,SQ1:C29,SQ1:C30,SQ1:C31,SQ1:C32,SQ1:C33,SQ1:C34,SQ1:C35,SQ1:C36,SQ1:C37,SQ1:C38)    
                IF clip(Mismo)='' OR clip(Mismo) <> clip(CAMPO)
                    Mismo = CAMPO    
                    SELF.QCampos.No = self.PosicionAgrupador
                    GET(SELF.QCampos,SELF.QCampos.No)
                    Case SELF.QCampos.Picture
                    OF 'D'
                        IF CLIP(CAMPO)<>''
                            ObjXls.ClearFormat()                
                            ObjXls.Format.Picture='@d06-'
                            ObjXls.Format.ExcelMask='dd/mm/yyyy'
                            ObjXls.Format.FontStyle=FONT:Bold
                            ObjXls.SetFormat()   
                            err#=ObjXls.WriteDateTime(REN#,COL#,DEFORMAT(CLIP(CAMPO),@d06),0);COL#+=1
                        ELSE
                            ObjXls.ClearFormat()                
                            err#=ObjXls.WriteString(REN#,COL#,'');COL#+=1    
                        END
                    OF 'A'
                        IF CLIP(CAMPO)<>''
                            ObjXls.ClearFormat()                
                            ObjXls.Format.Picture='@d06-'
                            
                            ObjXls.Format.ExcelMask='mm/dd/yyyy'
                            ObjXls.Format.FontStyle=FONT:Bold
                            ObjXls.SetFormat()   
                            err#=ObjXls.WriteDateTime(REN#,COL#,DEFORMAT(CLIP(CAMPO),@d06),0);COL#+=1
                        ELSE
                            ObjXls.ClearFormat()           
                            ObjXls.Format.FontStyle=FONT:Bold
                            ObjXls.SetFormat()   
                            err#=ObjXls.WriteString(REN#,COL#,'');COL#+=1    
                        END
                    
                    OF 'C'
                        ObjXls.ClearFormat()                
                        ObjXls.Format.ExcelMask='#,##0.00;-#,##0.00'    
                        ObjXls.Format.FontStyle=FONT:Bold
                        ObjXls.SetFormat()    
                        err#=ObjXls.WriteNumber(REN#,COL#,CLIP(CAMPO));COL#+=1
                        IF pTotaliza
                            SELF.QCampos.No = I#
                            GET(SELF.QCampos,SELF.QCampos.No)!SERIA LO MISMO QUE GET(SELF.QCampos,I#)
                            SELF.QCampos.Total+=CAMPO
                            PUT(SELF.QCampos)
                        END
                    OF 'N'   
                        ObjXls.ClearFormat()   
                        ObjXls.Format.FontStyle=FONT:Bold
                        ObjXls.SetFormat() 
                        err#=ObjXls.WriteNumber(REN#,COL#,CLIP(CAMPO));COL#+=1
                    OF 'S'
                        ObjXls.ClearFormat()                
                        ObjXls.Format.FontStyle=FONT:Bold
                        ObjXls.SetFormat() 
                        err#=ObjXls.WriteString(REN#,COL#,CLIP(CAMPO));COL#+=1
                    END
                    ren#+=1  
                END
            END
            
            ?PROGRESS999{PROP:Progress} = ?PROGRESS999{PROP:Progress} + 1
            COL#=1
            LOOP I#=1 TO SELF.ContCampos
                CAMPO = CHOOSE(I#,SQ1:C1,SQ1:C2,SQ1:C3,SQ1:C4,SQ1:C5,SQ1:C6,SQ1:C7,SQ1:C8,SQ1:C9,SQ1:C10,SQ1:C11,SQ1:C12,SQ1:C13,SQ1:C14,SQ1:C15,SQ1:C16,SQ1:C17,SQ1:C18,SQ1:C19,SQ1:C20,SQ1:C21,SQ1:C22,SQ1:C23,SQ1:C24,SQ1:C25,SQ1:C26,SQ1:C27,SQ1:C28,SQ1:C29,SQ1:C30,SQ1:C31,SQ1:C32,SQ1:C33,SQ1:C34,SQ1:C35,SQ1:C36,SQ1:C37,SQ1:C38)
                SELF.QCampos.No = I#
                GET(SELF.QCampos,SELF.QCampos.No)
                if ERRORCODE()  then cycle END
                Case SELF.QCampos.Picture
                OF 'D'
                    IF CLIP(CAMPO)<>''
                        ObjXls.ClearFormat()                
                        ObjXls.Format.Picture='@d06-'
                        ObjXls.Format.ExcelMask='dd/mm/yyyy'
                        ObjXls.SetFormat()   
                        err#=ObjXls.WriteDateTime(REN#,COL#,DEFORMAT(CLIP(CAMPO),@d06),0);COL#+=1
                    ELSE
                        ObjXls.ClearFormat()                
                        err#=ObjXls.WriteString(REN#,COL#,'');COL#+=1    
                    END
                OF 'A'
                    IF CLIP(CAMPO)<>''
                        ObjXls.ClearFormat()                
                        ObjXls.Format.Picture='@d06-'
                        ObjXls.Format.ExcelMask='mm/dd/yyyy'
                        ObjXls.SetFormat()   
                        err#=ObjXls.WriteDateTime(REN#,COL#,DEFORMAT(CLIP(CAMPO),@d06),0);COL#+=1
                    ELSE
                        ObjXls.ClearFormat()                
                        err#=ObjXls.WriteString(REN#,COL#,'');COL#+=1    
                    END
                    
                OF 'C'
                    ObjXls.ClearFormat()                
                    ObjXls.Format.ExcelMask='#,##0.00;-#,##0.00'      
                    ObjXls.SetFormat()    
                    err#=ObjXls.WriteNumber(REN#,COL#,CLIP(CAMPO));COL#+=1
                    IF pTotaliza
                        SELF.QCampos.No = I#
                        GET(SELF.QCampos,SELF.QCampos.No)!SERIA LO MISMO QUE GET(SELF.QCampos,I#)
                        SELF.QCampos.Total+=CAMPO
                        PUT(SELF.QCampos)
                    END
                OF 'N'   
                    ObjXls.ClearFormat()                
                    err#=ObjXls.WriteNumber(REN#,COL#,CLIP(CAMPO));COL#+=1
                OF 'S'
                    ObjXls.ClearFormat()                
                    err#=ObjXls.WriteString(REN#,COL#,CLIP(CAMPO));COL#+=1
                END
            END
            ren#+=1        
        END
        close(SQLFile)
        
        IF pTotaliza
            ObjXls.ClearFormat()
            ObjXls.Format.FontSize=12
            ObjXls.Format.FontStyle=FONT:Bold
            ObjXls.Format.ExcelMask='#,##0.00;-#,##0.00'      
            ObjXls.SetFormat()
            LOOP I#=1 TO RECORDS(SELF.QCampos)
                GET(SELF.QCampos,I#)
                IF ~SELF.QCampos.Total THEN CYCLE END
                err#=ObjXls.WriteNumber(REN#,I#,SELF.QCampos.Total)
            END
        END
        err#=ObjXls.Autofilter(3,1,ren#-1,COL#-1)
        ObjXls.FreezePanes(4,1)
        err#=ObjXls.CloseWorkbook()
        dispose(ObjXls)
        SELF.AbrirArchivo(0{prop:handle},SELF.FnOutPut)

		
        CLOSE(pgWindow)
!------------------------------------------------------------------------------------------
!ya implementado ....
!------------------------------------------------------------------------------------------
CreaExcelClassNew.GenerarReporteQ   PROCEDURE(STRING pTitulo,STRING pHeaders,BYTE pTotaliza=0)
CAMPO                                   CSTRING(200)
oStr                                    SystemStringClass
oStr2                                   SystemStringClass
ObjXls                                  &xlsxwriter 
i                                       LONG

pgWindow                                WINDOW,AT(,,306,28),FONT('Segoe UI',10),CENTER,COLOR(00C0FFFFh),GRAY
                                            STRING('Exportando a Excel por favor espere'),AT(2,2,303),USE(?STRING999),FONT(,,00CC6600h, |
                                                FONT:regular),CENTER
                                            PROGRESS,AT(7,14,297),USE(?PROGRESS999),RANGE(0,100)
                                        END

    CODE
        OPEN(pgWindow)
        DISPLAY
        
        ?PROGRESS999{PROP:RangeHigh}=RECORDS(SELF.QExport)
        !Descomponempos el query para obtener los campos del mismo
        FREE(SELF.QCampos)
        oStr.Str(clip(pHeaders))  
        oStr.Split(',')
        LOOP i=1 to  oStr.GetLinesCount()
            SELF.QCampos.Picture = 'S'
            SELF.QCampos.Nombre = oStr.GetLineValue(i)
            !buscamos si tiene un picture tal como [N]Numero,[D]Fecha,[S]String
            p#=INSTRING('[N]',oStr.GetLineValue(i),1,1)
            IF P#
                SELF.QCampos.Picture = 'N'
                SELF.QCampos.Nombre  = SELF.Remplazar(oStr.GetLineValue(i),'[N]','')
                !MESSAGE('Encontro picture=N')
            END
            p#=INSTRING('[D]',oStr.GetLineValue(i),1,1)
            IF P#
                SELF.QCampos.Picture = 'D'
                SELF.QCampos.Nombre  = SELF.Remplazar(oStr.GetLineValue(i),'[D]','')
                !MESSAGE('Encontro picture=D')
            END
            p#=INSTRING('[A]',oStr.GetLineValue(i),1,1)
            IF P#
                SELF.QCampos.Picture = 'A'
                SELF.QCampos.Nombre  = SELF.Remplazar(oStr.GetLineValue(i),'[A]','')
                !MESSAGE('Encontro picture=D')
            END
            p#=INSTRING('[S]',oStr.GetLineValue(i),1,1)
            IF P#
                SELF.QCampos.Picture = 'S'
                SELF.QCampos.Nombre  = SELF.Remplazar(oStr.GetLineValue(i),'[S]','')
                !MESSAGE('Encontro picture=S')
            END
            p#=INSTRING('[C]',oStr.GetLineValue(i),1,1)
            IF P#
                SELF.QCampos.Picture = 'C'
                SELF.QCampos.Nombre  = SELF.Remplazar(oStr.GetLineValue(i),'[C]','')
                !MESSAGE('Encontro picture=S')
            END
            SELF.QCampos.No     = i
            SELF.QCampos.Total  = 0
            ADD(SELF.QCampos)
            SELF.ContCampos+=1
        END

        col#=1;ren#=1
        ObjXls&=NEW(xlsxwriter) 
        ObjXls.NewWorkbook(SELF.FnOutPut)
        worksheet#=ObjXls.AddSheet('Sheet1')
        !-----------Parametros de la hoja-------------------
        col#=1;ren#=1
        ObjXls.ClearFormat()
        ObjXls.SetSelection(1,1,2,2)
        ObjXls.Format.FontSize=14
        ObjXls.Format.FontStyle=FONT:Bold
        ObjXls.SetFormat()
        !err#=ObjXls.WriteString(ren#,col#,clip(pTitulo))
        oStr2.Str(clip(pTitulo))
        ostr2.Split(',')
        LOOP i#=1 to  oStr2.GetLinesCount()
            z#=ObjXls.Merge(ren#,1,ren#,SELF.ContCampos)
            err#=ObjXls.WriteString(ren#,col#,clip(oStr2.GetLineValue(i#)));ren#+=1
        END
        ren#=3
        LOOP i#=1 TO SELF.ContCampos
            GET(SELF.QCampos,i#)
            if ~ERRORCODE()
                err#=ObjXls.WriteString(ren#,col#,CLIP(SELF.QCampos.Nombre));col#+=1
            END        
        END
        ren#=4

        LOOP ZZ#=1 TO RECORDS(SELF.QExport)
            GET(SELF.QExport,ZZ#)
            ?PROGRESS999{PROP:Progress} = ?PROGRESS999{PROP:Progress} + 1
            COL#=1
            LOOP I#=1 TO SELF.ContCampos
                CAMPO = CHOOSE(I#,SELF.QExport.QC1,SELF.QExport.QC2,SELF.QExport.QC3,SELF.QExport.QC4,SELF.QExport.QC5,SELF.QExport.QC6,SELF.QExport.QC7,SELF.QExport.QC8,SELF.QExport.QC9,SELF.QExport.QC10,SELF.QExport.QC11,SELF.QExport.QC12,SELF.QExport.QC13,SELF.QExport.QC14,SELF.QExport.QC15,SELF.QExport.QC16,SELF.QExport.QC17,SELF.QExport.QC18,SELF.QExport.QC19,SELF.QExport.QC20,SELF.QExport.QC21,SELF.QExport.QC22,SELF.QExport.QC23,SELF.QExport.QC24,SELF.QExport.QC25,SELF.QExport.QC26,SELF.QExport.QC27,SELF.QExport.QC28,SELF.QExport.QC29,SELF.QExport.QC30,SELF.QExport.QC31,SELF.QExport.QC32,SELF.QExport.QC33,SELF.QExport.QC34,SELF.QExport.QC35,SELF.QExport.QC36,SELF.QExport.QC37,SELF.QExport.QC38,SELF.QExport.QC39,SELF.QExport.QC40)
                SELF.QCampos.No = I#
                GET(SELF.QCampos,SELF.QCampos.No)
                if ERRORCODE()  then cycle END
                Case SELF.QCampos.Picture
                OF 'D'
                    IF CLIP(CAMPO)<>''
                        ObjXls.ClearFormat()                
                        ObjXls.Format.Picture='@d06-'
                        ObjXls.Format.ExcelMask='dd/mm/yyyy'
                        ObjXls.SetFormat()   
                        err#=ObjXls.WriteDateTime(REN#,COL#,DEFORMAT(CLIP(CAMPO),@d06),0);COL#+=1
                    ELSE
                        ObjXls.ClearFormat()                
                        err#=ObjXls.WriteString(REN#,COL#,'');COL#+=1    
                    END
                OF 'A'
                    IF CLIP(CAMPO)<>''
                        ObjXls.ClearFormat()                
                        ObjXls.Format.Picture='@d06-'
                        ObjXls.Format.ExcelMask='mm/dd/yyyy'
                        ObjXls.SetFormat()   
                        err#=ObjXls.WriteDateTime(REN#,COL#,DEFORMAT(CLIP(CAMPO),@d06),0);COL#+=1
                    ELSE
                        ObjXls.ClearFormat()                
                        err#=ObjXls.WriteString(REN#,COL#,'');COL#+=1    
                    END
                OF 'C'
                    ObjXls.ClearFormat()                
                    ObjXls.Format.ExcelMask='#,##0.00;-#,##0.00'      
                    ObjXls.SetFormat()    
                    err#=ObjXls.WriteNumber(REN#,COL#,CLIP(CAMPO));COL#+=1
                    IF pTotaliza
                        SELF.QCampos.No = I#
                        GET(SELF.QCampos,SELF.QCampos.No)!SERIA LO MISMO QUE GET(SELF.QCampos,I#)
                        SELF.QCampos.Total+=CAMPO
                        PUT(SELF.QCampos)
                    END
                OF 'S'
                    ObjXls.ClearFormat()                
                    err#=ObjXls.WriteString(REN#,COL#,CLIP(CAMPO));COL#+=1
                OF 'N'    
                    ObjXls.ClearFormat()                
                    err#=ObjXls.WriteNumber(REN#,COL#,CLIP(CAMPO));COL#+=1
!                    IF pTotaliza
!                        SELF.QCampos.No = I#
!                        GET(SELF.QCampos,SELF.QCampos.No)!SERIA LO MISMO QUE GET(SELF.QCampos,I#)
!                        SELF.QCampos.Total+=CAMPO
!                        PUT(SELF.QCampos)
!                    END
                END
            END
            ren#+=1   
        END!loop
        IF pTotaliza
            ObjXls.ClearFormat()
            ObjXls.Format.FontSize=12
            ObjXls.Format.FontStyle=FONT:Bold
            ObjXls.Format.ExcelMask='#,##0.00;-#,##0.00'      
            ObjXls.SetFormat()
            LOOP I#=1 TO RECORDS(SELF.QCampos)
                GET(SELF.QCampos,I#)
                IF ~SELF.QCampos.Total THEN CYCLE END
                err#=ObjXls.WriteNumber(REN#,I#,SELF.QCampos.Total)
            END
        END        
        
        err#=ObjXls.Autofilter(3,1,ren#-1,COL#-1)
        ObjXls.FreezePanes(4,1)
        err#=ObjXls.CloseWorkbook()
        dispose(ObjXls)
        SELF.AbrirArchivo(0{prop:handle},SELF.FnOutPut)

        CLOSE(pgWindow)
           
        
!---------------------------------------------------
! Descompone el query y llena los campos 
! de los encabezados al queue
!---------------------------------------------------
CreaExcelClassNew.LLenaCampos       PROCEDURE(STRING pHeader)
campos                                  STRING(3000)
oStr                                    SystemStringClass
i                                       ULONG
    CODE
        FREE(SELF.QCampos)
        x# = len(SELF.SQuery)
        p#=INSTRING('FROM',clip(self.SQuery),1,1)! buscamos la posicion del from 
        Campos = sub(clip(self.SQuery),7,P#-8)   !extraemos el string de los campos 
        self.SQCount=SELF.Remplazar(clip(SELF.SQuery),clip(campos),' COUNT(*) ') !generamos el query para el conteo de registros
        z#=INSTRING('GROUP BY',CLIP(self.SQCount),1,1)! le quitamos al count el group by si es que lo tiene 
        self.SQCount=SUB(self.SQCount,1,Z#-1)   
        z#=INSTRING('ORDER BY',CLIP(self.SQCount),1,1) ! y tambien el order by en caso de que lo tenga
        self.SQCount=SUB(self.SQCount,1,Z#-1)
        
        IF CLIP(pHeader)<>''! si se le pasaron los campos los seteamos
            campos=pHeader    
        END
        !con el objeto stringclass agregamos los campos a un queue
        oStr.Str(clip(Campos))  
        oStr.Split(',')
        LOOP i=1 to  oStr.GetLinesCount()
            SELF.QCampos.Picture = 'S'
            SELF.QCampos.Nombre = oStr.GetLineValue(i)
            if clip(self.NombreAgrupador)<>'' and clip(UPPER(self.NombreAgrupador)) = clip(UPPER(SELF.QCampos.Nombre))
                SELF.PosicionAgrupador = i
            END
            !buscamos si tiene un picture tal como [N]Numero,[D]Fecha,[S]String
            p#=INSTRING('[N]',oStr.GetLineValue(i),1,1)
            IF P#
                SELF.QCampos.Picture = 'N'
                SELF.QCampos.Nombre  = SELF.Remplazar(oStr.GetLineValue(i),'[N]','')
            END
            p#=INSTRING('[D]',oStr.GetLineValue(i),1,1)
            IF P#
                SELF.QCampos.Picture = 'D'
                SELF.QCampos.Nombre  = SELF.Remplazar(oStr.GetLineValue(i),'[D]','')
            END
            p#=INSTRING('[A]',oStr.GetLineValue(i),1,1)
            IF P#
                SELF.QCampos.Picture = 'A'
                SELF.QCampos.Nombre  = SELF.Remplazar(oStr.GetLineValue(i),'[A]','')
            END
            p#=INSTRING('[S]',oStr.GetLineValue(i),1,1)
            IF P#
                SELF.QCampos.Picture = 'S'
                SELF.QCampos.Nombre  = SELF.Remplazar(oStr.GetLineValue(i),'[S]','')
            END
            p#=INSTRING('[C]',oStr.GetLineValue(i),1,1)
            IF P#
                SELF.QCampos.Picture = 'C'
                SELF.QCampos.Nombre  = SELF.Remplazar(oStr.GetLineValue(i),'[C]','')
            END
            SELF.QCampos.No     = i
            SELF.QCampos.Total  = 0
            ADD(SELF.QCampos)
            SELF.ContCampos+=1
        END
!---------------------------------------------------------------        
!Remplaza uno o mas caracteres en un string 
!---------------------------------------------------------------        
CreaExcelClassNew.Remplazar PROCEDURE(STRING pValor,STRING pBuscar,STRING pRemplazar)!,STRING       
RetVal                          STRING(65520)
Pos                             LONG
Tl                              SHORT   !TokenLenght
    CODE
      
        RetVal = pValor
        Tl     = LEN(CLIP(pBuscar))
        IF RetVal = '' OR Tl = 0
            RETURN RetVal
        END
      
        Pos = INSTRING(pBuscar, RetVal, 1, 1)
        LOOP WHILE Pos
            RetVal = RetVal[1 : Pos-1] & CLIP(pRemplazar) & RetVal[Pos+Tl : 65520]
            Pos = INSTRING(pBuscar, RetVal, 1, 1)
        END
      
        RETURN RetVal        

        
!------------------------------------------------------------------        
!Implementacion del API ShellExecute
!------------------------------------------------------------------        
CreaExcelClassNew.AbrirArchivo      PROCEDURE(unsigned wHandle, STRING URL)        
URLBuffer                               CSTRING(256)
EmptyString                             CSTRING(254)
RetHandle                               LONG
RetMessage                              STRING(100)
    CODE                        
!        IF SELF.Abrirarchi=0 THEN RETURN END
        URLBuffer = CLIP(URL)
        EmptyString=''
        RetHandle=ShellExecuteZ(whandle, 0, URLBuffer, 0, EmptyString, 1)
        IF RetHandle =< 32
            CASE RetHandle
            OF 0
                RetMessage = 'Out of memory or file is corrupt when running program'
            OF 2
                RetMessage = 'File not found'
            OF 3
                RetMessage = 'Path not found'
            OF 5
                RetMessage = 'Sharing violation'
            OF 6
                RetMessage = 'Data segment error'
            OF 8
                RetMessage = 'Not enough memory to run program'
            OF 10
                RetMessage = 'Incorrect Windows version'
            OF 11
                RetMessage = 'Invalid program file.  Non Windows or corrupt .exe file'
            OF 12
                RetMessage = 'Not a Windows program'
            OF 13
                RetMessage = 'MS-DOS 4.0 program'
            OF 14
                RetMessage = 'Unknown program type'
            OF 15
                RetMessage = 'Can not run a real mode program'
            OF 16
                RetMessage = 'This program is already running and can only have one instance running'
            OF 19
                RetMessage = 'This is a compressed program file'
            OF 20
                RetMessage = 'One or more run time libraries are missing or corrupt'
            OF 21
                RetMessage = 'This program requires 32bit extension for Windows'
            OF 31
                RetMessage = 'No program associated with this file'
            Else
                RetMessage = ''
            END
            if RetMessage
                Message(Clip(Retmessage) & '<13><10>' & clip(URL))
            .
        END