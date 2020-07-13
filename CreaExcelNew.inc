PRAGMA('link(libxlsxw.lib)')
                    MAP
                        MODULE('ImpShell3')
                            ShellExecuteZ(UNSIGNED,LONG,*CSTRING,LONG,*CSTRING,SIGNED),UNSIGNED,PASCAL,RAW,NAME('SHELLEXECUTEA')
                        END

                    END

!Variables Globales
CamposQueue         QUEUE,TYPE,PRIVATE
Nombre                  STRING(100)
No                      BYTE
Total                   decimal(12,2)
Picture                 STRING(10)
                    END
QExportBase         QUEUE,TYPE
QC1                     CSTRING(201)
QC2                     CSTRING(201)
QC3                     CSTRING(201)
QC4                     CSTRING(201)
QC5                     CSTRING(201)
QC6                     CSTRING(201)
QC7                     CSTRING(201)
QC8                     CSTRING(201)
QC9                     CSTRING(201)
QC10                    CSTRING(201)
QC11                    CSTRING(201)
QC12                    CSTRING(201)
QC13                    CSTRING(201)
QC14                    CSTRING(201)
QC15                    CSTRING(201)
QC16                    CSTRING(201)
QC17                    CSTRING(201)
QC18                    CSTRING(201)
QC19                    CSTRING(201)
QC20                    CSTRING(201)
QC21                    CSTRING(201)
QC22                    CSTRING(201)
QC23                    CSTRING(201)
QC24                    CSTRING(201)
QC25                    CSTRING(201)
QC26                    CSTRING(201)
QC27                    CSTRING(201)
QC28                    CSTRING(201)
QC29                    CSTRING(201)
QC30                    CSTRING(201)
QC31                    CSTRING(201)
QC32                    CSTRING(201)
QC33                    CSTRING(201)
QC34                    CSTRING(201)
QC35                    CSTRING(201)
QC36                    CSTRING(201)
QC37                    CSTRING(201)
QC38                    CSTRING(201)
QC39                    CSTRING(201)
QC40                    CSTRING(201)
                    END



CreaExcelClassNew      class,module('CreaExcelNew.Clw'),link('CreaExcelNew')!,DLL(1),EXTERNAL 
!Propiedades
FnOutPut                STRING(250),PRIVATE   !archivo de salida
SQuery                  STRING(3000),PRIVATE  !query 
QCampos                 &CamposQueue,PRIVATE  !queue para almacenar los headers del reporte
FnOwnerName             STRING(200),PRIVATE   !conexion contra la db
ContCampos              LONG,PRIVATE          !conteo de headers
SQCount                 STRING(3000),PRIVATE  !query de tipo (select count(*) xxx)
TotalizaCampo           BYTE
Abrirarchi              BYTE
NombreAgrupador         CSTRING(20)
PosicionAgrupador       LONG

QExport                 &QExportBase


!--------------------------------------------------------------------------------------
!Metodos
!-------------------------------------------------------------------------------------
Construct               PROCEDURE()
Destruct                PROCEDURE()
Init                    PROCEDURE(STRING pFnOutput,STRING pOwner,byte Abrirarchi=1)
GenerarReporte          PROCEDURE(STRING pQuery,STRING pTitulo,STRING pHeaders,BYTE pTotaliza=0,<string pAgrupa>)
!agregado 23 jun 17 ----------------------------------------------------------------
GenerarReporteQ          PROCEDURE(STRING pTitulo,STRING pHeaders,BYTE pTotaliza=0)
!-----------------------------------------------------------------------------------
Remplazar               PROCEDURE(STRING pValor,STRING pBuscar,STRING pRemplazar),STRING 
LLenaCampos             PROCEDURE(STRING pHeader),PRIVATE
AbrirArchivo            PROCEDURE(unsigned wHandle, STRING URL)
                    END




