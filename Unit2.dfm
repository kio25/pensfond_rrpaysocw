object DataModule2: TDataModule2
  OldCreateOrder = False
  Left = 378
  Top = 413
  Height = 379
  Width = 776
  object OracleLogon1: TOracleLogon
    Session = OracleSession1
    Options = [ldAuto, ldDatabase, ldDatabaseList, ldLogonHistory, ldConnectAs]
    HistoryIniFile = 'c:\history.ini'
    Left = 696
    Top = 24
  end
  object OracleSession1: TOracleSession
    Left = 696
    Top = 88
  end
  object OracleDataSet1: TOracleDataSet
    SQL.Strings = (
      
        'SELECT EMP# emp,MONTH,NVL(SUMZPLMAX,0) sumzplmax ,NVL(SUMGPD,0) ' +
        'sumgpd ,NVL(SUMBOL,0) sumbol,'
      
        '       NVL(SUMBLFSS,0)sumblfss ,NVL(STAX_ZPL,0) stax_zpl, NVL(ST' +
        'AX_GPD,0) stax_gpd,'
      
        '       NVL(STAX_BL,0) STAX_BL,NVL(STAX_BLFSS,0) STAX_BLfss, NVL(' +
        'SUMZPL,0) sumzplp'
      '                 FROM PFTSUMN'
      '                 WHERE YEAR=:year1'
      '                       AND EMP# between :empmin AND :empmax'
      '                 ORDER BY EMP#,MONTH'
      ''
      ''
      ''
      '')
    Variables.Data = {
      0300000003000000070000003A454D504D494E03000000000000000000000007
      0000003A454D504D4158030000000000000000000000060000003A5945415231
      030000000000000000000000}
    QBEDefinition.QBEFieldDefs = {
      040000000B00000003000000454D50010000000000050000004D4F4E54480100
      000000000900000053554D5A504C4D41580100000000000600000053554D4750
      440100000000000600000053554D424F4C0100000000000800000053554D424C
      46535301000000000008000000535441585F5A504C0100000000000800000053
      5441585F47504401000000000007000000535441585F424C0100000000000A00
      0000535441585F424C4653530100000000000700000053554D5A504C50010000
      000000}
    Cursor = crSQLWait
    Session = OracleSession1
    Left = 40
    Top = 88
    object OracleDataSet1EMP: TIntegerField
      FieldName = 'EMP'
      Required = True
    end
    object OracleDataSet1MONTH: TIntegerField
      FieldName = 'MONTH'
      Required = True
    end
    object OracleDataSet1SUMZPLMAX: TFloatField
      FieldName = 'SUMZPLMAX'
    end
    object OracleDataSet1SUMGPD: TFloatField
      FieldName = 'SUMGPD'
    end
    object OracleDataSet1SUMBOL: TFloatField
      FieldName = 'SUMBOL'
    end
    object OracleDataSet1SUMBLFSS: TFloatField
      FieldName = 'SUMBLFSS'
    end
    object OracleDataSet1STAX_ZPL: TFloatField
      FieldName = 'STAX_ZPL'
    end
    object OracleDataSet1STAX_GPD: TFloatField
      FieldName = 'STAX_GPD'
    end
    object OracleDataSet1STAX_BL: TFloatField
      FieldName = 'STAX_BL'
    end
    object OracleDataSet1STAX_BLFSS: TFloatField
      FieldName = 'STAX_BLFSS'
    end
    object OracleDataSet1SUMZPLP: TFloatField
      FieldName = 'SUMZPLP'
    end
  end
  object OracleQuery1del: TOracleQuery
    SQL.Strings = (
      'DELETE FROM RECALC'
      '       WHERE YEARREC = :year1  AND'
      '                    YEAR=:year AND'
      '                    MONTH=:month AND'
      '                    PAY# IN ( 540,541,542,633 ) AND'
      '                    EMP# between :empmin and :empmax'
      ' ')
    Session = OracleSession1
    Variables.Data = {
      0300000005000000070000003A454D504D494E03000000000000000000000007
      0000003A454D504D4158030000000000000000000000060000003A4D4F4E5448
      030000000000000000000000050000003A594541520300000000000000000000
      00060000003A5945415231030000000000000000000000}
    Cursor = crSQLWait
    Left = 160
    Top = 16
  end
  object OracleQuery1ins: TOracleQuery
    SQL.Strings = (
      
        ' INSERT INTO RECALC(emp#,year,month,yearrec,monthrec,sum,pay#,ds' +
        'chdate,dbuser)'
      '          VALUES (:emp,:year,:month,:year1,:monthz,:RAZ,:pay,'
      
        '                   decode(length(ltrim(rtrim(:dschdate))),10,to_' +
        'date(:dschdate,'#39'dd/mm/yyyy'#39'),NULL)  ,'
      
        '                   to_char(SYSDATE,'#39'DD/MM/YYYY HH24:MI '#39') || USE' +
        'R)'
      '       ')
    Session = OracleSession1
    Variables.Data = {
      0300000008000000040000003A454D5003000000000000000000000005000000
      3A59454152030000000000000000000000060000003A4D4F4E54480300000000
      00000000000000060000003A5945415231030000000000000000000000070000
      003A4D4F4E54485A030000000000000000000000040000003A52415A04000000
      0000000000000000040000003A50415903000000000000000000000009000000
      3A4453434844415445050000000000000000000000}
    Cursor = crSQLWait
    Left = 152
    Top = 80
  end
  object OracleDataSetkoef: TOracleDataSet
    SQL.Strings = (
      'SELECT NVL(NVAL,0) koef,INDATE '
      #9' FROM   SYSINDEX                                     '
      #9' WHERE  IND#   = :kf    AND                         '
      #9'        ROWNUM = 1       AND                         '
      #9'        INDATE = (SELECT MAX(INDATE)                 '
      #9#9'        FROM   SYSINDEX                      '
      #9#9'        WHERE  IND# = :kf AND               '
      #9#9#9'     INDATE <= :date1)               '
      '')
    Variables.Data = {
      0300000002000000030000003A4B46030000000000000000000000060000003A
      44415445310C0000000000000000000000}
    QBEDefinition.QBEFieldDefs = {
      0400000002000000040000004B4F454601000000000006000000494E44415445
      010000000000}
    Cursor = crSQLWait
    Session = OracleSession1
    Left = 144
    Top = 152
    object OracleDataSetkoefKOEF: TFloatField
      FieldName = 'KOEF'
    end
    object OracleDataSetkoefINDATE: TDateTimeField
      FieldName = 'INDATE'
      Required = True
    end
  end
  object OracleQueryuvol: TOracleQuery
    SQL.Strings = (
      'SELECT  rtrim(NVL(to_char(DSCHDATE,'#39'dd/mm/yyyy'#39'),'#39' '#39')) dschdate'
      '        FROM EMPLOY'
      '        WHERE EMP# = :emp'
      '')
    Session = OracleSession1
    Variables.Data = {0300000001000000040000003A454D50030000000000000000000000}
    Cursor = crSQLWait
    Left = 272
    Top = 16
  end
  object OracleQuerysecret: TOracleQuery
    SQL.Strings = (
      'SELECT NVL(PRV#,0) pse'
      '               FROM EMPPRV'
      '              WHERE EMP#=:emp'
      '                AND PRV# in(998,999)'
      
        '                AND to_date('#39'01/'#39'||ltrim(to_char(:monthR,'#39'09'#39'))|' +
        '|'#39'/'#39'||ltrim(to_char(:yearR)),'#39'dd/mm/yyyy'#39')>=PRVDATE1'
      
        '                AND to_date('#39'01/'#39'||ltrim(to_char(:monthR,'#39'09'#39'))|' +
        '|'#39'/'#39'||ltrim(to_char(:yearR)),'#39'dd/mm/yyyy'#39')<NVL(PRVDATE2,to_date(' +
        #39'31/12/2999'#39','#39'dd/mm/yyyy'#39'))')
    Session = OracleSession1
    Variables.Data = {
      0300000003000000040000003A454D5003000000000000000000000007000000
      3A4D4F4E544852030000000000000000000000060000003A5945415252030000
      000000000000000000}
    Cursor = crSQLWait
    Left = 264
    Top = 104
  end
  object OracleDataSetsecret: TOracleDataSet
    SQL.Strings = (
      'sELECT NVL(PRV#,0) pse'
      '               FROM EMPPRV'
      '              WHERE EMP#=:emp'
      '                AND PRV# in(998,999)'
      
        '                AND to_date('#39'01/'#39'||ltrim(to_char(:monthR,'#39'09'#39'))|' +
        '|'#39'/'#39'||ltrim(to_char(:yearR)),'#39'dd/mm/yyyy'#39')>=PRVDATE1'
      
        '                AND to_date('#39'01/'#39'||ltrim(to_char(:monthR,'#39'09'#39'))|' +
        '|'#39'/'#39'||ltrim(to_char(:yearR)),'#39'dd/mm/yyyy'#39')<NVL(PRVDATE2,to_date(' +
        #39'31/12/2999'#39','#39'dd/mm/yyyy'#39'))'
      '')
    Variables.Data = {
      0300000003000000040000003A454D5003000000000000000000000007000000
      3A4D4F4E544852030000000000000000000000060000003A5945415252030000
      000000000000000000}
    QBEDefinition.QBEFieldDefs = {040000000100000003000000505345010000000000}
    Cursor = crSQLWait
    Session = OracleSession1
    Left = 360
    Top = 32
    object OracleDataSetsecretPSE: TFloatField
      FieldName = 'PSE'
    end
  end
  object OracleDataSetuvol: TOracleDataSet
    SQL.Strings = (
      'SELECT  rtrim(NVL(to_char(DSCHDATE,'#39'dd/mm/yyyy'#39'),'#39' '#39')) dschdate'
      '        FROM EMPLOY'
      '        WHERE EMP# = :emp'
      '')
    Variables.Data = {0300000001000000040000003A454D50030000000000000000000000}
    QBEDefinition.QBEFieldDefs = {0400000001000000080000004453434844415445010000000000}
    Session = OracleSession1
    Left = 136
    Top = 216
    object OracleDataSetuvolDSCHDATE: TStringField
      FieldName = 'DSCHDATE'
      Size = 75
    end
  end
end
