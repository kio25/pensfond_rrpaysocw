
unit Unit2;

interface

uses
  SysUtils, Classes, Oracle, DB, OracleData,Windows,Dialogs ,
  Messages,  Variants,  Controls, Math, ComObj,
  StdCtrls,  ComCtrls;
type
  TDataModule2 = class(TDataModule)
    OracleLogon1: TOracleLogon;
    OracleSession1: TOracleSession;
    OracleDataSet1: TOracleDataSet;
    OracleQuery1del: TOracleQuery;
    OracleQuery1ins: TOracleQuery;
    OracleDataSet1EMP: TIntegerField;
    OracleDataSet1MONTH: TIntegerField;
    OracleDataSet1SUMZPLMAX: TFloatField;
    OracleDataSet1SUMGPD: TFloatField;
    OracleDataSet1SUMBOL: TFloatField;
    OracleDataSet1SUMBLFSS: TFloatField;
    OracleDataSet1STAX_ZPL: TFloatField;
    OracleDataSet1STAX_GPD: TFloatField;
    OracleDataSet1STAX_BL: TFloatField;
    OracleDataSet1STAX_BLFSS: TFloatField;
    OracleDataSetkoef: TOracleDataSet;
    OracleDataSetkoefKOEF: TFloatField;
    OracleDataSetkoefINDATE: TDateTimeField;
    OracleQueryuvol: TOracleQuery;
    OracleQuerysecret: TOracleQuery;
    OracleDataSetsecret: TOracleDataSet;
    OracleDataSetsecretPSE: TFloatField;
    OracleDataSetuvol: TOracleDataSet;
    OracleDataSetuvolDSCHDATE: TStringField;
    OracleDataSet1SUMZPLP: TFloatField;
  private
    { Private declarations }
       procedure zap_baza(pay:integer;raz:real);
  public
   procedure raschet1;

 //  procedure raschet2;
    { Public declarations }
  end;

var
  DataModule2: TDataModule2;
  i,flaginv,emp,pay,MONTHz,k,ks,l,ls: integer;
  date1: TDateTime;
  summax,sum1,sumt1,sgpd,szplmax,szpl,st_blfss,raz :real;
  sbol,sblfss,sbolm,sblfssm,st_zpl,st_gpd,st_bl :real;
        dschdate: string;  //дата увольнени€
 //  myFile : TextFile;




implementation
 uses unit1;
{$R *.dfm}


function RoundEx(const AValue: Double; const ADigit: Integer = -2): Double;
var
  s:  String;
  st: Int64;
  sf: Real;
begin
if(AValue>=0) then begin
  s  := FloatToStr(AValue * IntPower(10, -ADigit));
  st := Trunc(StrToFloat(s));
  sf := Frac(StrToFloat(s));
  if sf <  0.5 then Result := st*IntPower(10, ADigit);
  if sf >= 0.5 then Result := (st+1)*IntPower(10, ADigit);
              end
              else
              begin
  s  := FloatToStr(AValue * IntPower(10, -ADigit));
  st := Trunc(StrToFloat(s));
  sf := Frac(StrToFloat(s));
  if sf >  -0.5 then Result := st*IntPower(10, ADigit);
  if sf <= -0.5 then Result := (st-1)*IntPower(10, ADigit);
              end;


//  if Temp >= 0.5 then ScaledFractPart := ScaledFractPart + 1;
//  if Temp <=-0.5 then ScaledFractPart := ScaledFractPart - 1;

end;


{/*

function RoundEx(X: Double; Precision: Integer): Double;
//Precision : 1 - до цел_х, 10 -до дес€тых, 100 - до сотых...
var
  ScaledFractPart, Temp : Double;
begin
  ScaledFractPart := Frac(X)*Precision;
  Temp := Frac(ScaledFractPart);
  ScaledFractPart := Int(ScaledFractPart);
  if Temp >= 0.5 then ScaledFractPart := ScaledFractPart + 1;
  if Temp <= -0.5 then ScaledFractPart := ScaledFractPart - 1;
  RoundEx := Int(X) + ScaledFractPart/Precision;
end;

}


procedure TDataModule2.raschet1;

   var
      emp_old,i,j,pse: integer;
      SumMaxR,SUMZPL,sumgpd,sumbol,sumblfss,stax_zpl,stax_gpd,stax_bl,stax_blfss,SUMZPLP:real;

      date1: TDateTime;
      mas_kf:array [1..5,1..12] of real;
      RUzpl,RAZzpl,RUgpd,RAZgpd,RUbol,RAZbol,RUblfss,RAZblfss:real;




    begin
 Form1.ProgressBar1.Min;
 if pr=0  //если указатель пустой то пишем в базу
  then  begin

                   OracleQuery1del.SetVariable('empmin',empmin);
                   OracleQuery1del.SetVariable('empmax',empmax);
                   OracleQuery1del.setVariable('month',month);
                   OracleQuery1del.SetVariable('year',year);
                   OracleQuery1del.SetVariable('year1',year1);

                   with OracleQuery1del do
                      try
                          Form1.StaticText1.Caption := 'del';
                       try
                          Execute;
                   except
                       on E:EOracleError do begin
                       ShowMessage(E.Message);
                       exit;
                       end;
                       end;

                   except
                       on E:EOracleError do ShowMessage(E.Message);
                   end;
               OracleSession1.Commit;
             end;  // if  pr=0



  OracleDataSet1.Close;
  OracleDataSet1.SetVariable('year1',year1);
  OracleDataSet1.SetVariable('empmin',empmin);
  OracleDataSet1.SetVariable('empmax',empmax);
  OracleDataSet1.Open;
  OracleDataSet1.First;


  if OracleDataSet1.RecordCount<>0
     then  begin
                   Form1.ProgressBar1.Max:=OracleDataSet1.RecordCount;
                   Form1.StaticText2.Caption:='0';
                   Form1.StaticText2.Repaint;
                   Form1.StaticText3.Caption:=Inttostr(OracleDataSet1.RecordCount);
                    Form1.StaticText3.Repaint;
{

  XLApp:=CreateOleObject('Excel.Application');
  XLApps:=CreateOleObject('Excel.Application');


  XlApp.Workbooks.Add(ExtractFilePath(ParamStr(0))+'rep540.xls');
  XlApps.Workbooks.Add(ExtractFilePath(ParamStr(0))+'rep540s.xls');
  XLApp.Workbooks[1].Worksheets[1].Name:='540';
  XLApps.Workbooks[1].Worksheets[1].Name:='540s';



  Colum:=XLApp.Workbooks[1].WorkSheets['540'].Columns;
  Row:=XLApp.Workbooks[1].WorkSheets['540'].Rows;
  Sheet:=XLApp.Workbooks[1].WorkSheets['540'];
  Sheet.Cells[2,5]:=IntToStr(year1)+' год.';

  Colums:=XLApps.Workbooks[1].WorkSheets['540s'].Columns;
  Rows:=XLApps.Workbooks[1].WorkSheets['540s'].Rows;
  Sheets:=XLApps.Workbooks[1].WorkSheets['540s'];
  Sheets.Cells[2,5]:=IntToStr(year1)+' год.';

  k:=10;ks:=10;   l:=0;ls:=0;

 }
  //вычисление коэф.

    for i:=1 to 12 do
      for j:=1 to 5 do
          mas_kf[j,i]:=0;


    for i:=1 to 12 do
    begin
             //  form1.ProgressBar1.Step := 50 div month2;
                 date1:=StrToDate('01.'+IntToStr(i)+'.'+IntToStr(year1));
                 OracleDataSetkoef.Close;
                 OracleDataSetkoef.SetVariable('date1',date1);
                 OracleDataSetkoef.SetVariable('kf',541);   //summax
                 OracleDataSetkoef.Open;
                 OracleDataSetkoef.First;
                 mas_kf[1,i]:= OracleDataSetkoefKOEF.AsFloat;

                 OracleDataSetkoef.Close;
                 OracleDataSetkoef.SetVariable('date1',date1);
                 OracleDataSetkoef.SetVariable('kf',563);      //kfzpl
                 OracleDataSetkoef.Open;
                 OracleDataSetkoef.First;
                 mas_kf[2,i]:= OracleDataSetkoefKOEF.AsFloat;

                 OracleDataSetkoef.Close;
                 OracleDataSetkoef.SetVariable('date1',date1);
                 OracleDataSetkoef.SetVariable('kf',566);      //kfgpd
                 OracleDataSetkoef.Open;
                 OracleDataSetkoef.First;
                 mas_kf[3,i]:= OracleDataSetkoefKOEF.AsFloat;

                 OracleDataSetkoef.Close;
                 OracleDataSetkoef.SetVariable('date1',date1);
                 OracleDataSetkoef.SetVariable('kf',564);        //kfbol
                 OracleDataSetkoef.Open;
                 OracleDataSetkoef.First;
                 mas_kf[4,i]:= OracleDataSetkoefKOEF.AsFloat;


                 OracleDataSetkoef.Close;
                 OracleDataSetkoef.SetVariable('date1',date1);
                 OracleDataSetkoef.SetVariable('kf',565);      //kfblfss
                 OracleDataSetkoef.Open;
                 OracleDataSetkoef.First;
                 mas_kf[5,i]:= OracleDataSetkoefKOEF.AsFloat;


      {     //печать коэф.
                write(myfile,inttostr(i)+' ');
                 for j:=1 to 5 do
                 write(myfile,inttostr(j)+' '+floattostr(mas_kf[j,i])+' ');
                 writeln(myfile,'  ');
       }
               //  ProgressBar1.StepIt;
    end ;
              OracleDataSetkoef.Close;
      emp_old:=0;
      pse:=0;
   for i:=1 to OracleDataSet1.RecordCount do
     begin

                Form1.ProgressBar1.Position:=i;
                  Form1.StaticText2.Caption:=Inttostr(i);
                  Form1.StaticText2.Repaint;

      emp:=0;  monthz:=0;
      SUMZPL:=0;  SUMGPD:=0; SUMBOL:=0;  SUMBLFSS:=0;  SUMzPLP:=0;
      STAX_ZPL:=0;  STAX_GPD:=0;  STAX_BL:=0;  STAX_BLFSS:=0;



      RUzpl:=0;  RAZzpl:=0; RUgpd:=0; RAZgpd:=0;
      RUbol:=0;   RAZbol:=0;  RUblfss:=0; RAZblfss:=0;

      emp:=OracleDataSet1EMP.AsInteger;
      monthz:=OracleDataSet1MONTH.AsInteger;
      SUMZPL:=OracleDataSet1SUMZPLMAX.AsFloat;
      SUMGPD:=OracleDataSet1SUMGPD.AsFloat;
      SUMBOL:=OracleDataSet1SUMBOL.AsFloat;
      SUMBLFSS:=OracleDataSet1SUMBLFSS.AsFloat;
      STAX_ZPL:=OracleDataSet1STAX_ZPL.AsFloat;
      STAX_GPD:=OracleDataSet1STAX_GPD.AsFloat;
      STAX_BL:=OracleDataSet1STAX_BL.AsFloat;
      STAX_BLFSS:=OracleDataSet1STAX_BLFSS.AsFloat;
      sumzplP:=OracleDataSet1SUMZPLp.AsFloat;


                 if emp_old<>emp then  begin
                             OracleDataSetuvol.Close;
                 OracleDataSetuvol.SetVariable('emp',emp);
                 OracleDataSetuvol.Open;
                 OracleDataSetuvol.First;
                   if OracleDataSetuvol.RecordCount<>0
                        then  dschdate:=OracleDataSetuvolDSCHDATE.AsString
                        else  dschdate:=' ';     //дл€ таб. который нет в EMPLOY



{                 OracleQueryuvol.SetVariable('emp',emp);
                   with OracleQueryuvol do
                      try
                          Form1.StaticText1.Caption := 'uvol '+inttostr(emp);
                       try
                          Execute;
                   except
                       on E:EOracleError do begin
                       ShowMessage(E.Message);

                       exit;
                       end;
                       end;
                       dschdate:=Field(0);
                   except
                       on E:EOracleError do  ShowMessage(E.Message);
                   end;
}
                 OracleDataSetsecret.Close;
                 OracleDataSetsecret.SetVariable('emp',emp);
                 OracleDataSetsecret.SetVariable('monthr',month);
                 OracleDataSetsecret.SetVariable('yearr',year);
                 OracleDataSetsecret.Open;
                 OracleDataSetsecret.First;
                   if OracleDataSetsecret.RecordCount<>0
                        then pse:=OracleDataSetsecretPSE.AsInteger
                        else pse:=0;



                                  end; // if emp_old<>emp then

      //////////////////
         if ((mas_kf[1,monthz]-SUMZPL) <= 0 )  //if1 ((sumMax - SUMZPL) <= 0 )
           then begin
             RUzpl:=RoundEx(mas_kf[1,monthz]*mas_kf[2,monthz],-2);         // sumMax * KFzpl
             RAZzpl:=RUzpl-STAX_ZPL;
             RAZbol:=RUbol-STAX_BL;
             RAZblfss:=RUblfss-STAX_BLFSS;
        //   ѕереход на 4
               end
           else begin
                RUzpl:=RoundEx(SUMZPL*mas_kf[2,monthz],-2);         // SUMZPL*KFzpl
                RAZzpl:=RUzpl-STAX_ZPL;
                SumMaxR:=mas_kf[1,monthz]-SUMZPL ;

                if ((SumMaxR-SUMBOL)<= 0)  //if2
                     then begin
                          RUbol:=RoundEx(SumMaxR*mas_kf[4,monthz],-2);       //SumMaxR*KFbol
                          RAZbol:=RUbol-STAX_BL;
                          RAZblfss:=RUblfss-STAX_BLFSS;
                          //    ѕереход на 4
                         end

                     else  begin
                           RUbol:=RoundEx(SUMBOL*mas_kf[4,monthz],-2);    //SUMBOL*KFbol
                           RAZbol:=RUbol-STAX_BL;
                           SumMaxR:=SumMaxR-SUMBOL;

                          if ((SumMaxR-SUMBLFSS )<= 0 )   //if3
                              then  begin
                                  RUblfss:=RoundEx(SumMaxR*mas_kf[5,monthz],-2);  //SumMaxR * KFblfss
                                  RAZblfss:=RUblfss-STAX_BLFSS;
                                  // ѕереход на 4
                                  end
                              else   begin
                                      RUblfss:=RoundEx(SUMBLFSS*mas_kf[5,monthz],-2);   //SUMBLFSS * KFblfss
                                      RAZblfss:=RUblfss-STAX_BLFSS;
                                      end;                               //if3
               end ;  //if2 else
               end; //if1 else

      if pr=0  //если указатель пустой то пишем в базу
           then  begin
              if RoundEx(Razzpl,-2)<>0
                 then  zap_baza(540,RoundEx(Razzpl,-2));//  пишем в базу   pay=540

              if RoundEx(Razbol,-2)<>0
                 then zap_baza(541,RoundEx(Razbol,-2)); // пишем в базу   pay=541

              if RoundEx(RAZblfss,-2)<>0
                 then zap_baza(633,RoundEx(RAZblfss,-2)); // пишем в базу   pay=633

             if ((RoundEx(Razzpl,-2)<>0) or (RoundEx(Razbol,-2)<>0) or (RoundEx(RAZblfss,-2)<>0))
                          then     OracleSession1.Commit;
                end;  //      if pr=0


   if pse=0 then begin
        if (emp_old<>emp)  then    begin
           emp_old:=emp;  k:=k+1;
                                    end;
           Sheet.Cells[k,1]:=emp;
       if dschdate<>'' then  Sheet.Cells[k,9]:=dschdate;
//          k:=k+1;
//     writeln(myfile,inttostr(emp)+' дата увольнени€ '+dschdate);
//     writeln(myfile,inttostr(monthz)+' 1 '+floattostr(SUMZPL)+' '+floattostr(mas_kf[1,monthz])+' '+floattostr(RUzpl)+' '+floattostr(STAX_ZPL)+' '+floattostr(RoundEx(RAZzpl,-2)));
        Sheet.Cells[k,2]:=monthz; Sheet.Cells[k,3]:='1';Sheet.Cells[k,4]:=SUMZPLp;
        Sheet.Cells[k,5]:=mas_kf[1,monthz]; Sheet.Cells[k,6]:=RUzpl  ;Sheet.Cells[k,7]:=STAX_ZPL;
        Sheet.Cells[k,8]:=RoundEx(RAZzpl,-2); k:=k+1;

     if ( (SUMBOL<>0) or (STAX_BL<>0) or (RUbol<>0) or (RAZbol<>0) ) then begin
         Sheet.Cells[k,1]:=emp;
        Sheet.Cells[k,2]:=monthz; Sheet.Cells[k,3]:='2';Sheet.Cells[k,4]:=SUMBOL;
        Sheet.Cells[k,5]:=mas_kf[1,monthz]; Sheet.Cells[k,6]:=RUbol  ;Sheet.Cells[k,7]:=STAX_BL;
        Sheet.Cells[k,8]:=RoundEx(RAZbol,-2); k:=k+1;
//        writeln(myfile,inttostr(monthz)+' 2 '+floattostr(SUMBOL)+' '+floattostr(mas_kf[1,monthz])+' '+floattostr(RUbol)+' '+floattostr(STAX_BL)+' '+floattostr(RoundEx(RAZbol,-2)));
                                                                           end;
     if ( (SUMBLFSS<>0) or (STAX_BLFSS<>0) or (RUblfss<>0) or (RAZblfss<>0) ) then  begin
        Sheet.Cells[k,1]:=emp;
        Sheet.Cells[k,2]:=monthz; Sheet.Cells[k,3]:='3';Sheet.Cells[k,4]:=SUMBLFSS;
        Sheet.Cells[k,5]:=mas_kf[1,monthz]; Sheet.Cells[k,6]:=RUblfss  ;Sheet.Cells[k,7]:=STAX_BLFSS;
        Sheet.Cells[k,8]:=RoundEx(RAZblfss,-2); k:=k+1;
//                writeln(myfile,inttostr(monthz)+' 3 '+floattostr(SUMBLFSS)+' '+floattostr(mas_kf[1,monthz])+' '+floattostr(RUblfss)+' '+floattostr(STAX_BLFSS)+' '+floattostr(RoundEx(RAZblfss,-2)));
                                                                         END;
     if ( (SUMGPD<>0) or (STAX_GPD<>0) or (RUgpd<>0) or (RAZgpd<>0) ) then   begin Sheet.Cells[k,1]:=emp;
        Sheet.Cells[k,2]:=monthz; Sheet.Cells[k,3]:='4';Sheet.Cells[k,4]:=SUMGPD;
        Sheet.Cells[k,5]:=mas_kf[1,monthz]; Sheet.Cells[k,6]:=RUGPD  ;Sheet.Cells[k,7]:=STAX_GPD;
        Sheet.Cells[k,8]:=RoundEx(RAZGPD,-2); k:=k+1;
//        writeln(myfile,inttostr(monthz)+' 4 '+floattostr(SUMGPD)+' '+floattostr(mas_kf[1,monthz])+' '+floattostr(RUgpd)+' '+floattostr(STAX_GPD)+' '+floattostr(RoundEx(RAZgpd,-2)));
                                                                         END;
                  end   // if pse=0 then
       else   begin
        if (emp_old<>emp)  then
           begin emp_old:=emp; ks:=ks+1;
//                 writeln(myfiles,inttostr(emp)+' дата увольнени€ '+dschdate);
               end;
//     writeln(myfiles,inttostr(monthz)+' 1 '+floattostr(SUMZPL)+' '+floattostr(mas_kf[1,monthz])+' '+floattostr(RUzpl)+' '+floattostr(STAX_ZPL)+' '+floattostr(RAZzpl));
           Sheets.Cells[ks,1]:=emp;
       if dschdate<>'' then  Sheets.Cells[ks,9]:=dschdate;
        Sheets.Cells[ks,2]:=monthz; Sheets.Cells[ks,3]:='1';Sheets.Cells[ks,4]:=SUMZPLp;
        Sheets.Cells[ks,5]:=mas_kf[1,monthz]; Sheets.Cells[ks,6]:=RUzpl  ;Sheets.Cells[ks,7]:=STAX_ZPL;
        Sheets.Cells[ks,8]:=RoundEx(RAZzpl,-2); ks:=ks+1;


     if ( (SUMBOL<>0) or (STAX_BL<>0) or (RUbol<>0) or (RAZbol<>0) ) then begin
         Sheets.Cells[ks,1]:=emp;
        Sheets.Cells[ks,2]:=monthz; Sheets.Cells[ks,3]:='2';Sheets.Cells[ks,4]:=SUMBOL;
        Sheets.Cells[ks,5]:=mas_kf[1,monthz]; Sheets.Cells[ks,6]:=RUbol  ;Sheets.Cells[ks,7]:=STAX_BL;
        Sheets.Cells[ks,8]:=RoundEx(RAZbol,-2); ks:=ks+1;

//        writeln(myfiles,inttostr(monthz)+' 2 '+floattostr(SUMBOL)+' '+floattostr(mas_kf[1,monthz])+' '+floattostr(RUbol)+' '+floattostr(STAX_BL)+' '+floattostr(RAZbol));
                                                    end;
     if ( (SUMBLFSS<>0) or (STAX_BLFSS<>0) or (RUblfss<>0) or (RAZblfss<>0) ) then begin
        Sheets.Cells[ks,1]:=emp;
        Sheets.Cells[ks,2]:=monthz; Sheets.Cells[ks,3]:='3';Sheets.Cells[ks,4]:=SUMBLFSS;
        Sheets.Cells[ks,5]:=mas_kf[1,monthz]; Sheets.Cells[ks,6]:=RUblfss  ;Sheets.Cells[ks,7]:=STAX_BLFSS;
        Sheets.Cells[ks,8]:=RoundEx(RAZblfss,-2); ks:=ks+1;

     //        writeln(myfiles,inttostr(monthz)+' 3 '+floattostr(SUMBLFSS)+' '+floattostr(mas_kf[1,monthz])+' '+floattostr(RUblfss)+' '+floattostr(STAX_BLFSS)+' '+floattostr(RAZblfss));
              end;
     if ( (SUMGPD<>0) or (STAX_GPD<>0) or (RUgpd<>0) or (RAZgpd<>0) ) then begin
        Sheets.Cells[ks,1]:=emp;
        Sheets.Cells[ks,2]:=monthz; Sheets.Cells[ks,3]:='4';Sheets.Cells[ks,4]:=SUMGPD;
        Sheets.Cells[ks,5]:=mas_kf[1,monthz]; Sheets.Cells[ks,6]:=RUGPD  ;Sheets.Cells[ks,7]:=STAX_GPD;
        Sheets.Cells[ks,8]:=RoundEx(RAZGPD,-2); ks:=ks+1;

     //        writeln(myfiles,inttostr(monthz)+' 4 '+floattostr(SUMGPD)+' '+floattostr(mas_kf[1,monthz])+' '+floattostr(RUgpd)+' '+floattostr(STAX_GPD)+' '+floattostr(RAZgpd));
            end;
                  end;  // if pse=0 else

     //если записей больше чем макс кол-во строк в Excel
      if (k>=50000) then begin
                 inc(l);
                 XLApp.Workbooks[1].Worksheets[l+1].Name:='540_'+inttostr(l);
                 Colum:=XLApp.Workbooks[1].WorkSheets['540_'+inttostr(l)].Columns;
                 Row:=XLApp.Workbooks[1].WorkSheets['540_'+inttostr(l)].Rows;
                 Sheet:=XLApp.Workbooks[1].WorkSheets['540_'+inttostr(l)];
                 k:=1;
                      end;

      if (ks>=50000) then begin
           inc(ls);
           XLApps.Workbooks[1].Worksheets[l+1].Name:='540s_'+inttostr(ls);
           Colums:=XLApps.Workbooks[1].WorkSheets['540s_'+inttostr(ls)].Columns;
           Rows:=XLApps.Workbooks[1].WorkSheets['540s_'+inttostr(ls)].Rows;
           Sheets:=XLApps.Workbooks[1].WorkSheets['540s_'+inttostr(ls)];
           ks:=1;
                     end;


      OracleDataSet1.Next;
      ////////////////////

     end; //    for i:=1 to OracleDataSet1.RecordCount do





   {
    vers:=VarToStr(XLApp.version);
    office:='';
    FileName:='';
    office:=Copy(vers,1,Pos('.',vers)-1);

     if length(trim(form1.Edit2.Text))=1
       then  mmf:='0'+ trim(form1.Edit2.Text)
       else  mmf:=trim(form1.Edit2.Text);

     if StrToInt(office)>11 then
      begin
       FileName:='c:\rep540_'+trim(form1.Edit3.Text)+'_'+mmf+'.xlsx';
       FileNames:='c:\rep540s_'+trim(form1.Edit3.Text)+'_'+mmf+'.xlsx';

      end
       else
      begin
       FileName:='c:\rep540_'+trim(form1.Edit3.Text)+'_'+mmf+'.xls';
       FileNames:='c:\rep540_'+trim(form1.Edit3.Text)+'_'+mmf+'.xls';

      end;

     XLApp.Workbooks[1].SaveAs(FileName);
     XLApps.Workbooks[1].SaveAs(FileNames);

     XLApp.Quit;  XLApp:=Unassigned;
     XLApps.Quit;  XLApps:=Unassigned;
    }

//     MyExcel.WorkBooks.Item[WBIndex].SaveAs(FileName);
              // XLApps.Visible:=true;
              //  XLApp.Visible:=true;
           end   //if OracleDataSet1.RecordCount<>0 then begin
     else   pr_zap:=1; // showmessage('Ќет данных по запросу.');

    OracleDataSet1.Close;
    end;



 procedure  TDataModule2.zap_baza(pay:integer;raz:real);     //запись в базу
     begin

                   OracleQuery1ins.SetVariable('emp',emp);
                   OracleQuery1ins.SetVariable('year',year);
                   OracleQuery1ins.SetVariable('month',month);
                   OracleQuery1ins.SetVariable('year1',year1);
                   OracleQuery1ins.SetVariable('monthz',monthz);
                   OracleQuery1ins.SetVariable('raz',raz);
                   OracleQuery1ins.SetVariable('pay',pay);
                   OracleQuery1ins.SetVariable('dschdate',dschdate);
                     with OracleQuery1ins do
                      try
                          Form1.StaticText1.Caption := 'ins '+IntToStr(emp);
                       try
                          Execute;
                   except
                       on E:EOracleError do begin
                       ShowMessage(E.Message);
                       exit;
                       end;
                       end;

                   except
                       on E:EOracleError do ShowMessage(E.Message);
                   end;
  end; //procedure  TDataModule2.zap_baza;     //запись в базу



end.
