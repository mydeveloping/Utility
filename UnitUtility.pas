unit UnitUtility;
{$IFDEF FPC}
{$codepage utf8}
{$EndIf}

interface
  uses
    registry, Windows, SysUtils, Classes, ComCtrls, ComObj, Variants,
    stdctrls, Graphics, Controls, Forms, DB, RichEdit, Grids, Dialogs
    {$IFDEF FPC}, MaskEdit, LazUTF8,SynEdit{$EndIf}
    {$IFNDEF FPC}, mask, DBGridEh{$ENDIF}
    {$IFNDEF FPC}
    {$IFDEF MACOS} ,MacApi.Appkit,Macapi.CoreFoundation, Macapi.Foundation {$ENDIF}
    {$IFDEF MSWINDOWS} , Winapi.Messages {$ENDIF}
    {$ENDIF};

const

  {$REGION 'Mobile Constants'}
   // Use for Iranian cell phone number
    Begin0    = $00C080FF; // شماره با صفر شروع نشده
    Begin9    = $004080FF; // شماره با صفر شروع شده ولی 9 ندارد
    Begin09   = $000000FF; // شماره با 09 شروع نشده
    NumLes11  = clSilver ; // شماره کمتر از 11 رقم است
    NumBig11  = $00FF8080; // شماره بیشتر از 11 رقم است
    NumNotDig = clYellow ; // شماره شامل غیر رقم است
    NumTrue   = clTeal   ; // شماره درست است

  {$ENDREGION}

  {$REGION 'Natinal Code Constants'}
    // Use for national code number in Iran
    Valid             = 1; // شماره ملی درست است
    SameNumber        = 2; // شماره ملی فقط یک عدد مشابه است
    NumberLessThan10  = 4; // شماره ملی کمتر از 10 رقم دارد
    NumberBigerThan10 = 8; // شماره ملی بیشتر از 10 رقم دارد

  {$ENDREGION}

  {$REGION 'Numbers'}
    // Use for convert Arabic, English and Persian digits to anothers
    ArabicDigits  = #1632#1633#1634#1635#1636#1637#1638#1639#1640#1641;//'٠١٢٣٤٥٦٧٨٩';
    EnglishDigits = '0123456789';
    PersianDigits = #1776#1777#1778#1779#1780#1781#1782#1783#1784#1785;//'۰۱۲۳۴۵۶۷۸۹';

  {$ENDREGION}
  {$REGION 'Different Alphabet Between Arabic and Persian'}

    PersianChars = #1740#1705;// 'یک';
    ArabicChars  = #1610#1603;// 'يك';

  {$ENDREGION}

type

  TAlignType = (altLeft, altTop, altRight, altBottom, altVerticallyCenter,
               altHorizontallyCenter, altCenter, altHorizontallySpaceEqually,
               altVerticallySpaceEqually);
  TSizeType  = (stParentWidth, stParentHeight, stParentSize, stParentHalfWidth,
               stParentHalfHeight, stParentHalfsize, stParentOneThirdWidth,
               stParentOneThirdHeight, stParentOneThirdSize, stParentQuarterWidth,
               stParentQuarterHeight, stParentQuarterSize);

  SetOfAlignType = set of TAlignType;

  { TAlign }

  TAlign = class
    public
      class procedure AlignControl(AControl: TControl; AlignType: SetOfAlignType);
      class procedure AlignControls(AControlsContainer: TControl;
        Controls: array of TControl; AlignType: SetOfAlignType);overload;
      class procedure AlignControls(AControlsContainer: TWinControl;
        AlignType: SetOfAlignType);overload;
  end;

  TSizing = class
    public
      class procedure ResizeControl(AControl: TControl; SizeType: TSizeType);overload;
      class procedure ResizeControl(AControl: TControl; AWidth, AHeight: Integer); overload;
      // Grid columns resizing
      class procedure ResizeGridColumnsWidth(AGrid: TDrawGrid; ColumnsWidth: Word);overload;
      class procedure ResizeGridColumnsWidth(AGrid: TDrawGrid; ColumnsWidth: array of Word);overload;
      // Grid rows resizing
      class procedure ResizeGridRowsHeight(AGrid: TDrawGrid; RowsHeight: Word);overload;
      class procedure ResizeGridRowsHeight(AGrid: TDrawGrid; RowsHeight: array of Word);overload;
  end;

  TMemoUtility = class
    public
      class procedure SortLines(AMemo: TMemo);
      class procedure AddFilesToEnd(AMemo:TMemo;FileNames: TStrings);overload;
      class procedure AddFilesToEnd(AMemo:TMemo;FileName: string);overload;
      class procedure AddNumberToMemo(AMemo: TMemo; AnEditNumber:TEdit);
  end;

  {$IFNDEF FPC}
  TRichEditUtility = class
    // this procedure need RichEdit unit
    class procedure HighlightRichEditLine(ARichEdit: TRichEdit;ALine: Integer;
      BackColor: TColor; ForeColor: TColor = clBlack);
  end;

  TEditHelper = class Helper for TEdit
  private
    // Retrun raw number from separated number
    // if Edit text be 12,345 this function return 12345
    function GetNumber: string;
  public
    property Number: string read GetNumber;
  end;
  {$ENDIF}

  TStringUtility = class
    public
      // تبدیل < ک > و < ی > عربی به معادل فارسی آنها
      // Convert ک and ی from Persian form to Arabic form
      class function KafYaPersianToArabic(AText:string):String;
      // تبدیل < ک > و < ی > فارسی به معادل عربی آنها
      // Convert ک and ی from Arabic form to Persian form
      class function KafYaArabicToPersian(AText:string):String;
      // تشخیص اینکه آیا تمام نویسه های رشته عدد است یا خیر
      class function IsAllTextDigit(AText: string):Boolean;
      // تبدیل اعداد عربی داخل متن
      // به انگلیسی
      // Convert Arabic digits to English form
      class function ArabicDigitsToEnglish(AText: string): string;
      // به فارسی
      // Convert Arabic digits to Persian form
      class function ArabicDigitsToPersian(AText: string): string;
      // تبدیل اعداد انگلیسی داخل متن
      // به عربی
      // Convert English digits to Arabic form
      class function EnglishDigitsToArabic(AText: string): string;
      // به فارسی
      // Convert English digits to Persian form
      class function EnglishDigitsToPersian(AText: string): string;
      // تبدیل اعداد فارسی داخل متن
      // به انگلیسی
      // Convert Persian digits to English form
      class function PersianDigitsToEnglish(AText: string): string;
      // به عربی
      // Convert Persian digits to Arabic form
      class function PersianDigitsToArabic(AText: string): string;
      // وارونه کردن یک رشته
      // Reverse a string
      class function ReverseString(const AText: string): string;
  end;

  TNumberUtility = class
    // بدست آوردن تعداد ارقام عدد
    class function GetDigitNumber(ANumber: Int64): Byte;
    // نوشتن معادل عدد به حروف
    class function NumberToWordsPersian(ANumber: Int64): string;
    // Thanks to Paul Jackson
    // at http://www.delphicode.co.uk/number-to-words-function/
    class function NumberToWordsEnglish(ANumber: Int64): string;
    // سه رقم جدا کردن عدد و بازگرداندن آن
    class function ThosandsSeparate(ANumber: string)  : string;overload;
    class function ThosandsSeparate(ANumber: Extended): string;overload;
    class function ThosandsSeparate(ANumber: Int64)   : string;overload;
  end;

  TMobileUtility = class
    // Check form of any cell phone number in Persian phone number form
    // بررسی درستی یک شماره تلفن همراه
    class function CheckNumberForm(ANumber: string): TColor;
  end;

  TNationalCode = class
    // اعتبارسنجی شماره ملی افراد در ایران
    // Validating of national code of Iranian persons.
    class function ValidateNationalCode(NationalCode: string): Byte;
    // اعتبارسنجی شماره ملی اشخاص حقوقی و شرکت ها
    // Validating of national code of Legal entity of Iranian companies.
    class function ValidateLegalEntityNationalCode(LENC: string): boolean;
  end;

  TDBUtility = class
    class function TrueFieldValue(AValue: string;
      AFieldType: TfieldType = ftString):string;
  end;

  { TWinControlUtility }

  TWinControlUtility = class
    class procedure SetBiDiToRightToLeft(AWinControl: TWinControl);
    class procedure ClearAllControlValues(AWinControl: TWinControl);
  end;

  TKeyboardUtility = class
    class procedure ChangeKeyboardToFarsi;
    class procedure ChangeKeyboardToEnglish;
  end;

  TApplicationUtility = class
    // بدست آوردن اطلاعات فایل اجرایی پروژه در زمانی که در حال اجرا است.
    class procedure GetBuildInfo(var V1, V2, V3, V4: word);
    // بدست آوردن اطلاعات فایل اجرایی به صورت رشته ای
    class function GetBuildInfoAsString: string;
  end;
  {$IfNDef FPC}
  TStringHelper = record Helper for string
  private
    function GetNumber: string;
  public
    property Number: string read GetNumber ;
  end;
  {$EndIf}

  {$IFNDEF FPC}
  TExport = class
    private
//      function RefToCell(ARow, ACol: Integer): string;
    public
    // خروجی گرفتن از یک دی بی گرید در یک فایل
    // Export dbgrid values to a file
    class procedure DBGridToExcelFile(const AFileName: TFileName;
      const ADBGridEh: TDBGridEh);
    end;
  {$ENDIF}

  TFontUtility = class
    {$IFNDEF FPC}
    // Thanks to mehmed.ali
    // User at stackoverflow by this address
    // https://stackoverflow.com/users/605656/mehmed-ali
    class function CollectFonts: TStrings;
    {$ENDIF}
  end;

  function CreateRegistryKey(Root: HKEY; Path, Key: string):string;

implementation

function RefToCell(ARow, ACol: Integer): string;
begin
  Result := Chr(Ord('A') + ACol - 1) + IntToStr(ARow);
end;

class procedure TApplicationUtility.GetBuildInfo(var V1, V2, V3, V4: word);
var
  VerInfoSize, VerValueSize, Dummy: DWORD;
  VerInfo: Pointer;
  VerValue: PVSFixedFileInfo;
begin
  VerInfoSize := GetFileVersionInfoSize(PChar(ParamStr(0)), Dummy);
  if VerInfoSize > 0 then
  begin
      GetMem(VerInfo, VerInfoSize);
      try
        if GetFileVersionInfo(PChar(ParamStr(0)), 0, VerInfoSize, VerInfo) then
        begin
          VerQueryValue(VerInfo, '\', Pointer(VerValue), VerValueSize);
          with VerValue^ do
          begin
            V1 := dwFileVersionMS shr 16;
            V2 := dwFileVersionMS and $FFFF;
            V3 := dwFileVersionLS shr 16;
            V4 := dwFileVersionLS and $FFFF;
          end;
        end;
      finally
        FreeMem(VerInfo, VerInfoSize);
      end;
  end;
end;

class function TApplicationUtility.GetBuildInfoAsString: string;
var
  V1, V2, V3, V4: word;
begin
  GetBuildInfo(V1, V2, V3, V4);
  Result := IntToStr(V1) + '.' + IntToStr(V2) + '.' +
    IntToStr(V3) + '.' + IntToStr(V4);
end;


function CreateRegistryKey(Root: HKEY; Path, Key: string):string;
{$IfnDef FPC}
const
  RootPath: array [HKEY_CLASSES_ROOT..HKEY_DYN_DATA] of string =
    ('HKEY_CLASSES_ROOT','HKEY_CURRENT_USER','HKEY_LOCAL_MACHINE','HKEY_USERS',
    'HKEY_PERFORMANCE_DATA','HKEY_CURRENT_CONFIG','HKEY_DYN_DATA');
{$endif}
var
  Reg: TRegistry;
  Res, Exist: Boolean;
  TrimPath: string;
begin
  TrimPath := Trim(Path);
  if TrimPath = '' then
    Exit('There isn''t any path to create');
  if (Root = HKEY_LOCAL_MACHINE) and (UpperCase(copy(TrimPath,1,8)) <> 'SOFTWARE') then
    Exit('In HKEY_LOCAL_MACHINE key don''t create any key out of "SOFTWARE" super key');

  Reg := TRegistry.Create(KEY_READ);
  Reg.RootKey := HKEY_LOCAL_MACHINE;
  if (Root >= HKEY_CLASSES_ROOT) and (Root <= HKEY_DYN_DATA) then
    Reg.RootKey := Root;
  try
    try
      Exist := Reg.KeyExists(TrimPath+Key);
      if Exist then
        Result := 'Duplicate Value'
      else
      begin
        reg.Access:=KEY_WRITE;
        Res := reg.CreateKey(TrimPath+Key);
        {$ifndef fpc}if Res then
          Result := RootPath[Root] + '\' + TrimPath + '\' + Key
        else        {$endif}
          Result := 'Error';
      end;
    except
      on e: exception do
      begin
        Exit('Exception: ' + #10 + e.Message);
      end;
    end;
  finally
    reg.CloseKey;
    reg.Free;
  end;
end;

{ TMobile }

class function TMobileUtility.CheckNumberForm(ANumber: string): TColor;
begin
  // اگر در بین رشته نویسه ای غیر از ارقام باشند رنگ زرد
  if not TStringUtility.IsAllTextDigit(ANumber) then
    Result := NumNotDig
  // اگر عدد با 09 شروع نشود رنگ قرمز
  else if Copy(ANumber,1, 2) <> '09' then
  begin
    if ANumber[1] = '9' then
      Result := Begin0
    else if (ANumber[1] = '0') and (ANumber[2] <> '9') then
      Result := Begin9
    else
      Result := Begin09;
  end
  // اگر عدد کمتر از 11 رقم باشد رنگ خاکستری
  else if Length(ANumber) < 11 then
    Result := NumLes11
  // اگر عدد بیشتر از 11 رقم باشد رنگ ارغوانی
  else if Length(ANumber) > 11 then
    Result := NumBig11

//  else if not IsAllTextDigit(ANumber) then
//    Result := clBlue
  // و در صورت صحیح بودن شماره رنگ سبز
  else
    Result := NumTrue
end;

{ TString }

class function TStringUtility.ArabicDigitsToEnglish(AText: string): string;
var
  I: Integer;
  TempString: string;
begin
  TempString := AText;
  for I := 1 to Length(EnglishDigits) do
    TempString := StringReplace(TempString, ArabicDigits[i], EnglishDigits[i],
      [rfReplaceAll]);
  Result := TempString;
end;

class function TStringUtility.ArabicDigitsToPersian(AText: string): string;
var
  I: Integer;
  TempString: string;
begin
  TempString := AText;
  for I := 1 to Length(EnglishDigits) do
    TempString := StringReplace(TempString, ArabicDigits[i], PersianDigits[i],
      [rfReplaceAll]);
  Result := TempString;
end;

class function TStringUtility.EnglishDigitsToArabic(AText: string): string;
var
  I: Integer;
  TempString: string;
begin
  TempString := AText;
  for I := 1 to Length(EnglishDigits) do
    TempString := StringReplace(TempString, EnglishDigits[i], ArabicDigits[i],
      [rfReplaceAll]);
  Result := TempString;
end;

class function TStringUtility.EnglishDigitsToPersian(AText: string): string;
var
  I: Integer;
  TempString: string;
begin
  TempString := AText;
  for I := 1 to Length(EnglishDigits) do
    TempString := StringReplace(TempString, EnglishDigits[i], PersianDigits[i],
      [rfReplaceAll]);
  Result := TempString;
end;

class function TStringUtility.IsAllTextDigit(AText: string): Boolean;
var
  Temp: Int64;
begin
  // اگر در تبدیل رشته به عدد استثنائی رخ ندهد آن رشته تماما عدد است
  // در غیر اینصورت نویسه ای به جز ارقام هم در میان رشته وجود دارد
  // When convert a string to integer has not any exception
  // thats mean this string is a number
  Result := TryStrToInt64(Trim(AText), Temp);
end;

{ TMemoUtility }

class procedure TMemoUtility.AddFilesToEnd(AMemo:TMemo;FileNames: TStrings);
var
  I: Integer;
begin
  //      Unitutility.CreateregistryKey(HKEY_CURRENT_USER, );
  // حلقه ای برای باز کردن فایل ها و افزودن به شماره ها
  for I := 0 to FileNames.Count - 1 do
  begin
    AddFilesToEnd(AMemo, FileNames[i]);
  end;
end;

class procedure TMemoUtility.AddNumberToMemo(AMemo: TMemo; AnEditNumber: TEdit);
var
  msgres: integer;// مقدار برگشتی کادر محاوره ای کاربر
  msg: string; // پیام نمایش به کاربر
  I: Integer;  // شمارنده حلقه
  Duplicate: Boolean;// متغیری برای تشخیص تکراری بودن شماره ها
  ParentForm: TControl;
begin
  Duplicate := False;
  while not (ParentForm is tform) do
    ParentForm := ParentForm.Parent;
  // در صورتی که شماره صحیح باشد به این گزینش ارسال می گردد
  if AnEditNumber.Color = NumTrue then
  begin
    // در این حلقه تکراری بودن شماره بررسی می گردد
    for I := 0 to AMemo.Lines.Count - 1 do
      if AnEditNumber.Text = Trim(AMemo.Lines[i]) then
      begin
        Duplicate := True;
        Break;
      end;
    // اگر شماره تکراری باشد با نمایش پیام به کاربر تعامل او با این پیام را
    // دریافت می کند
    if Duplicate then
    begin
      msgres := MessageBox(TForm(ParentForm).Handle,
        'شماره تکراری است.' + #10 + 'آیا می خواهید شماره رونویسی شود؟',
        'خطا در اضافه کردن شماره'
        , MB_YESNO or MB_RIGHT or MB_RTLREADING or MB_ICONWARNING);
      if msgres = mrYes then
      begin
        AnEditNumber.Text := '';
        AnEditNumber.Color := clWindow;
      end
    end
    // اگر پیام تکراری نباشد آنرا به لیست می افزاید
    else
    begin
      AMemo.Lines.Add(AnEditNumber.Text);
      AnEditNumber.Text := '';
      AnEditNumber.Color := clWindow;
    end;
    AnEditNumber.SetFocus;
  end
  // اگر شماره صحیح نباشد با نمایش پیغام خطای مربوطه به کاربر اطلاع می دهد.
  else
  begin
    msg := 'خطا:' + #10;
    case AnEditNumber.Color of
      Begin09:
        msg := msg + 'شماره موبایل با "09" شروع می گردد.';
      NumLes11:
        msg := msg + 'شماره وارد شده کمتر از 11 رقم است.';
      NumBig11:
        msg := msg + 'شماره بیشتر از 11 رقم است.';
      NumNotDig:
        msg := msg + 'شماره وارد شده حاوی نویسه های غیرعددی می باشد.';
      Begin0:
        msg := msg + 'شماره ی وارد شده باید با 0 شروع شود.';
      Begin9:
        msg := msg + 'شماره ی وارد شده در رقم دوم یک 9 کم دارد.';
    end;
    msgres := MessageBox(TForm(ParentForm).Handle,
      pchar('شماره وارد شده صحیح نیست.' + #10 + msg + #10 +
      'آیا می خواهید شماره را ویرایش کنید؟'), 'خطا در ورود اطلاعات',
      MB_YESNO or MB_RIGHT or MB_RTLREADING or MB_ICONHAND);
    if msgres = mrYes then
      AnEditNumber.SetFocus;
  end;
end;

class procedure TMemoUtility.AddFilesToEnd(AMemo:TMemo;FileName: string);
var
  Temp: TStrings;
begin
  // ایجاد شی ممو
  Temp := TStringList.Create;
  try
    //      Unitutility.CreateregistryKey(HKEY_CURRENT_USER, );
    // حلقه ای برای باز کردن فایل ها و افزودن به شماره ها
//    for I := 0 to FileNames.Count - 1 do
    begin
      Temp.LoadFromFile(FileName);
      AMemo.Lines.AddStrings(Temp);
    end;
  finally
    Temp.Free;
  end;
end;

class procedure TMemoUtility.SortLines(AMemo: TMemo);
var
  Temp: TStringList;
begin
//  Result := nil;
  Temp := TStringList.Create;
  try
  Temp.AddStrings(AMemo.Lines);
  Temp.Sort;
  AMemo.Lines.Clear;
//  AMemo.Lines := Temp.Strings;
  AMemo.Lines.AddStrings(Temp);
  finally
    Temp.Free;
  end;
//  Result := Temp.Strings;
end;

{ TNationalCode }

class function TNationalCode.ValidateLegalEntityNationalCode(
  LENC: string): boolean;
var
  Len,
  ControlCode,
  Factor,
  Sum,
  Remaining : Integer;
  Flag : Boolean;
begin
  Result := false;
  Len := Length(LENC);
  Flag := (LENC = '00000000000') or (LENC = '11111111111') or (LENC = '22222222222') or (LENC = '33333333333');
  Flag := (LENC = '44444444444') or (LENC = '55555555555') or (LENC = '66666666666') or (LENC = '77777777777') or Flag;
  Flag := (LENC = '88888888888') or (LENC = '99999999999') or Flag;
  if (Flag = true) or (Len < 11) then
    Exit;
  ControlCode := StrToInt(LENC[11]);
  Factor := StrToInt(LENC[10]) + 2;
  Sum := 0;
  Sum := Sum + ((Factor + StrToInt(LENC[1])) * 29);
  Sum := Sum + ((Factor + StrToInt(LENC[2])) * 27);
  Sum := Sum + ((Factor + StrToInt(LENC[3])) * 23);
  Sum := Sum + ((Factor + StrToInt(LENC[4])) * 19);
  Sum := Sum + ((Factor + StrToInt(LENC[5])) * 17);
  Sum := Sum + ((Factor + StrToInt(LENC[6])) * 29);
  Sum := Sum + ((Factor + StrToInt(LENC[7])) * 27);
  Sum := Sum + ((Factor + StrToInt(LENC[8])) * 23);
  Sum := Sum + ((Factor + StrToInt(LENC[9])) * 19);
  Sum := Sum + ((Factor + StrToInt(LENC[10])) * 17);
  Remaining := Sum mod 11;
  if Remaining = 10 then
    Remaining := 0;
  Result := (Remaining = ControlCode);
end;

class function TNationalCode.ValidateNationalCode(NationalCode: string): Byte;
var
  I, Sum, a, c: Integer;
  First5, Last5, First2, First34: string;
begin
  Result := 0;
  if Length(NationalCode) < 10 then
    Exit(NumberLessThan10)
  else if Length(NationalCode) > 10 then
    Exit(NumberBigerThan10);

  Last5 := Copy(NationalCode, 5, 5);
  First5 := Copy(NationalCode, 1, 5);
  First2 := Copy(First5, 1, 2);
  First34:= Copy(First5, 3, 2);
  if (First5 = Last5)and(First2 = First34)and(First2[1] = First2[2])
    and(First2[1]= First5[3])then
    Exit(SameNumber);
  Sum := 0;
  for I := 1 to Length(NationalCode) - 1 do
    Sum := Sum + (11 - i) * StrToInt(NationalCode[i]);
  a := StrToInt(NationalCode[10]);
  c := sum mod 11;
  if ((c < 2) and (c = a))or
     ((c >= 2) and (a = 11 - c))then
    Result := Valid;
end;

{ TDBUtility }

// تابعی که مقادیر ارسال شده به آن را به فرم صحیح برای درج در بانک بر می گرداند
class function TDBUtility.TrueFieldValue(AValue: string;
  AFieldType: TfieldType): string;
begin
  // تنظیم به خالی برای فیلدهایی که خالی باشند
  // Default set null
  Result := 'NULL';
  // انتخابی که بسته به نوع فیلد مقدار صحیح را بر می گرداند
  // Case that return correct field value for query depend on field type
  case AFieldType of
    // ftUnknown: ;
    ftString, ftWideString:
      if Trim(AValue) <> '' then
        Result := QuotedStr(Trim(AValue));
    ftSmallint, ftInteger, ftWord
      ,{$ifndef fpc}ftByte, ftLongWord, ftShortint, {$endif}ftLargeint:
      if Trim(AValue) <> '' then
      begin
        try
          StrToInt64(Trim(AValue));
          Result := Trim(AValue);
        except
          Result := 'NULL';
        end;
      end;
    ftBoolean:
      if Trim(AValue) <> '' then
      begin
        try
          StrToBool(Trim(AValue));
          Result := Trim(AValue);
        except
          Result := 'NULL';
        end;
      end;
    ftFloat{$ifndef fpc}, ftSingle, ftExtended{$endif}:
      if Trim(AValue) <> '' then
      begin
        try
          StrToFloat(Trim(AValue));
          Result := Trim(AValue);
        except
          Result := 'NULL';
        end;
      end;
    // ftCurrency: ;
    // ftBCD: ;
    ftDate, ftTime, ftDateTime:
      if Trim(AValue) <> '' then
      begin
        try
          StrToDate(Trim(AValue));
          Result := QuotedStr(Trim(AValue));
        except
          Result := 'NULL';
        end;
      end;
    // ftBytes: ;
    // ftVarBytes: ;
    // ftAutoInc: ;
    // ftBlob: ;
    // ftMemo: ;
    // ftGraphic: ;
    // ftFmtMemo: ;
    // ftParadoxOle: ;
    // ftDBaseOle: ;
    // ftTypedBinary: ;
    // ftCursor: ;
    // ftFixedChar: ;
    // ftADT: ;
    // ftArray: ;
    // ftReference: ;
    // ftDataSet: ;
    // ftOraBlob: ;
    // ftOraClob: ;
    // ftVariant: ;
    // ftInterface: ;
    // ftIDispatch: ;
    // ftGuid: ;
    // ftTimeStamp: ;
    // ftFMTBcd: ;
    // ftFixedWideChar: ;
    // ftWideMemo: ;
    // ftOraTimeStamp: ;
    // ftOraInterval: ;
    // ftConnection: ;
    // ftParams: ;
    // ftStream: ;
    // ftTimeStampOffset: ;
    // ftObject: ;
  end;
end;

{ TFarsi }

class function TStringUtility.KafYaArabicToPersian(AText: string): string;
var
  TempString: {$IFDEF fpc} UTF8String {$ELse} string {$Endif};
  I: Integer;
begin
  {$IFDEF fpc}
    TempString := UTF8Decode(AText);
     for i := 1 to Length(TempString) do
       if TempString[i] = ArabicChars[1] then
         TempString[i] := PersianChars[1]
       else if TempString[i] = ArabicChars[2] then
         TempString[i] := PersianChars[2];
  {$ELse}
    TempString := AText;
    for I := 1 to Length(PersianChars) do
      TempString :=
        StringReplace(TempString, ArabicChars[i], PersianChars[i], [rfReplaceAll]);
  {$Endif}
  for i := 1 to Length(TempString) do
  Result := TempString;
end;

class function TStringUtility.KafYaPersianToArabic(AText: string): String;
var
  TempString: {$IFDEF FPC} UTF8String {$ELSE} string {$ENDIF};
  I: Integer;
begin
  {$IFDEF FPC}
    TempString := UTF8Decode(AText);
     for i := 1 to Length(TempString) do
       if TempString[i] = PersianChars[1] then
         TempString[i] := ArabicChars[1]
       else if TempString[i] = PersianChars[2] then
         TempString[i] := ArabicChars[2];
  {$ELSE}
    TempString := AText;
    for I := 1 to Length(ArabicChars) do
      TempString :=
        StringReplace(TempString, PersianChars[i], ArabicChars[i], [rfReplaceAll]);
  {$ENDIF}
  for i := 1 to Length(TempString) do
  Result := TempString;
end;

{ TWinControlUtility }

class procedure TWinControlUtility.SetBiDiToRightToLeft(AWinControl: TWinControl);
var
  ExStyle: Longint;
begin
  ExStyle := GetWindowLong(AWinControl.Handle, GWL_EXSTYLE);
  SetWindowLong(AWinControl.Handle, GWL_EXSTYLE, ExStyle or WS_EX_RTLREADING or
    WS_EX_RIGHT or WS_EX_LAYOUTRTL or WS_EX_NOINHERITLAYOUT);
  AWinControl.Repaint;
end;

class procedure TWinControlUtility.ClearAllControlValues(
  AWinControl: TWinControl);
var
  i: integer;
  TempControl: TControl;
  StringGridColCount: integer;
begin
  for i := 0 to AWinControl.ControlCount - 1 do
  begin
    TempControl := AWinControl.Controls[i];
    // برای زمانیکه یک کنترل خود نوعی پنجره است
    // Recursive call if control is a window control
    if TempControl is TWinControl then
      ClearAllControlValues(TWinControl(TempControl));
    // در زمان های معمولی
    if TempControl is TEdit then
      TEdit(TempControl).Clear;
    if TempControl is TComboBox then
      TComboBox(TempControl).ItemIndex := 0;
    if TempControl is TListBox then
      TListBox(TempControl).Items.Clear;
    if TempControl is TMaskEdit then
      TMaskEdit(TempControl).Clear;
    if TempControl is TStringGrid then
    begin
      StringGridColCount := TStringGrid(TempControl).ColCount;
      {$IFDEF FPC}
      TStringGrid(TempControl).Clear;
      TStringGrid(TempControl).RowCount := 2;
      TStringGrid(TempControl).ColCount := StringGridColCount;
      {$ELSE}
         TStringGrid(TempControl).RowCount := 2;
         TStringGrid(TempControl).Rows[1].Clear;
      {$ENDIF}
    end;
  end;
end;

{ TKeyboardUtility }

class procedure TKeyboardUtility.ChangeKeyboardToEnglish;
begin
  // تنظیم صفحه کلید به انگلیسی
  // Change keyboard to English
  LoadKeyboardLayout('LoadKeyboardLayout', 1);
end;

class procedure TKeyboardUtility.ChangeKeyboardToFarsi;
begin
  // تنظیم صفحه کلید به فارسی
  // change keyboard to Persian
  LoadKeyboardLayout('00000429', 1);
end;

{ TRichEditUtility }

{$IFNDEF FPC}
class procedure TRichEditUtility.HighlightRichEditLine(ARichEdit: TRichEdit;
  ALine: Integer; BackColor, ForeColor: TColor);
var
  // Temp variable for line value holder
  StrLine: string; // متغیر کمکی برای نگه داری مقدار یک خط
  // Format holder
  cf: _CHARFORMAT2W;// متغیری برای تغییر فرمت در ریج ادیت
begin
  // پر کردن متغیر با مقدار صفر
  // Fill cf variable with zero
  FillChar(cf, sizeof(cf), 0);
  // مقداردهی به اندازه
  // Initial cb size
  cf.cbSize := sizeof(cf);
  // مقدار دهی به رنگ پس زمینه
  // Intial background color
  cf.dwMask := CFM_BACKCOLOR;
  // مقداردهی به رنگ فونت
  // Initial font color
  cf.crBackColor := BackColor;
  // انتخاب یک خط از ریچ ادیت
  // Select a line of richedit
  StrLine := ARichEdit.Lines[ALine];
  // حذف خط انتخابی برای دوباره نویسی با فرمت دلخواه
  // Delete selected line for rewrite this
  ARichEdit.Lines.Delete(ALine);
  //  RichEditCorrection.SelAttributes.Color := Random(ColorToRGB(clWhite));
  // اعمال تغییرات در فرمت
  // Commit format change
  ARichEdit.Perform(EM_SETCHARFORMAT, SCF_SELECTION, lparam(@cf));
  // اعمال تغییر در رنگ فونت
  // Commit color change
  ARichEdit.SelAttributes.Color := ForeColor;
//  ARichEdit.SelAttributes.Style := RE.SelAttributes.Style + Style;
  // بازنویسی خط با فرمت جدید
  // Rewrite with new format
  ARichEdit.Lines.Insert(ALine, StrLine);
end;

{$ENDIF}

{ TNumberUtility }

class function TNumberUtility.GetDigitNumber(ANumber: Int64): Byte;
var
  TempInteger: Int64;
begin
  Result := 0;
  if ANumber <> 0 then
  begin
    TempInteger := ANumber;
    repeat
      TempInteger := TempInteger div 10;
      Inc(Result);
    until TempInteger=0;
  end;
end;

class function TStringUtility.PersianDigitsToArabic(AText: string): string;
var
  I: Integer;
  TempString: string;
begin
  TempString := AText;
  for I := 1 to Length(EnglishDigits) do
    TempString := StringReplace(TempString, PersianDigits[i], ArabicDigits[i],
      [rfReplaceAll]);
  Result := TempString;
end;

class function TStringUtility.PersianDigitsToEnglish(AText: string): string;
var
  I: Integer;
  TempString: string;
begin
  TempString := AText;
  for I := 1 to Length(EnglishDigits) do
    TempString := StringReplace(TempString, PersianDigits[i], EnglishDigits[i],
      [rfReplaceAll]);
  Result := TempString;
end;

class function TStringUtility.ReverseString(const AText: string): string;
var
  i, len: Integer;
begin
  len := Length(AText);
  SetLength(Result, len);
  for i := len downto 1 do
  begin
    Result[len - i + 1] := AText[i];
  end;
end;

class function TNumberUtility.ThosandsSeparate(ANumber: string): string;
var
  TempInt64: Int64;
  TempExtended: Extended;
begin
  if TryStrToInt64(ANumber, TempInt64) then
    Result := ThosandsSeparate(TempInt64)
  else if TryStrToFloat(ANumber, TempExtended) then
    Result := ThosandsSeparate(TempExtended)
  else
    Result := ANumber;

  if Result = '' then
    Result := '0';
end;

class function TNumberUtility.ThosandsSeparate(ANumber: Extended): string;
begin
  Result := FormatFloat('#,###,###.###', ANumber);
  if Result = '' then
    Result := '0';
end;

class function TNumberUtility.ThosandsSeparate(ANumber: Int64): string;
begin
  Result := FormatCurr('#,###,###', ANumber);
  if Result = '' then
    Result := '0';
end;

{ TExport }

{$IFNDEF FPC}
class procedure TExport.DBGridToExcelFile(const AFileName: TFileName;
  const ADBGridEh: TDBGridEh);
const
  Worksheet = -4167;
var
  Row, Col: Integer;
  Excel, Sheet, Data: OLEVariant;
  I, J, DataCols, DataRows: Integer;
begin

//  DataCols := ABSQuery1.FieldCount;
//  DataRows := ABSQuery1.RecordCount + 1; //1 for the title
  with ADBGridEh.DataSource.DataSet do
  begin
    DataCols := FieldCount;
    DataRows := RecordCount + 1; //1 for the title
  end;

  //Create a variant array the size of your data
  Data := VarArrayCreate([1, DataRows, 1, DataCols], varVariant);

  //write the titles
  for I := 0 to DataCols - 1 do
    Data[1, I+1] := ADBGridEh.Columns[I].Title.Caption;

  //write data
  J := 1;
  with ADBGridEh.DataSource.DataSet do
  begin
    First;
    while (not Eof) and (J < DataRows) do
    begin
      for I := 0 to DataCols - 1 do
        Data[J + 1, I + 1] := Fields[I].Value;
      Inc(J);
      Next;
    end;
  end;

  //Create Excel-OLE Object
  Excel := CreateOleObject('Excel.Application');
  try
    //Don't show excel
    Excel.Visible := False;

    Excel.Workbooks.Add(Worksheet);
    Sheet := Excel.Workbooks[1].WorkSheets[1];
    Sheet.Name := 'Sheet1';
    //Fill up the sheet
    Sheet.Range[RefToCell(1, 1), RefToCell(DataRows, DataCols)].Value := Data;
    //Save Excel Worksheet
    try
      Excel.Workbooks[1].SaveAs(AFileName);
    except
      on E: Exception do
        raise Exception.Create('Data transfer error: ' + E.Message);
    end;
  finally
    if not VarIsEmpty(Excel) then
    begin
      Excel.DisplayAlerts := False;
      Excel.Quit;
      Excel := Unassigned;
      Sheet := Unassigned;
    end;
  end;
end;

{$ENDIF}

class function TNumberUtility.NumberToWordsPersian(ANumber: Int64): string;
const
  Hezar   = 1000;
  Milion  = Hezar * Hezar;
  Miliard = Milion * Hezar;

  NumberWords1: array [0..19] of string =
  ('صفر','یک','دو','سه','چهار','پنج','شش','هفت','هشت','نه', 'ده',
   'یازده', 'دوازده', 'سیزده','چهارده','پانزده','شانزده','هفده','هجده','نوزده');
  NumberWords10: array [2..9] of string =
  ('بیست', 'سی', 'چهل', 'پنجاه', 'شصت', 'هفتاد', 'هشتاد', 'نود');
  NumberWords100: array [1..9] of string =
  ('یکصد','دویست','سیصد','جهارصد','پانصد','ششصد','هفتصد','هشتصد','نهصد');
//  NumberWordsMul1000: array [1..3] of string =
//  ('هزار','میلیون','میلیارد');
begin
  if (ANumber < 0) then
    Result := 'منفی '
  else
    Result := '';

  ANumber := Abs(ANumber);

  if (ANumber >= 0) and (ANumber <= 19) then
    Result := NumberWords1[ANumber]
  else if (ANumber >= 20) and (ANumber <= 99) then
    begin
      if ANumber mod 10 = 0 then
        Result := NumberWords10[ANumber div 10]
      else
        Result := NumberWords10[ANumber div 10] + ' و '
          + NumberToWordsPersian(ANumber mod 10);
    end
  else if (ANumber >= 100) and (ANumber <= 999) then
    begin
      if ANumber mod 100 = 0 then
        Result := NumberWords100[ANumber div 100]
      else
        Result := NumberWords100[ANumber div 100] + ' و '
          + NumberToWordsPersian(ANumber mod 100);
    end
  else if (ANumber >= 1 * Hezar) and (ANumber < 1 * Milion) then
    begin
      if ANumber mod Hezar = 0 then
        Result := NumberToWordsPersian(ANumber div Hezar) + ' هزار'
      else
        Result := NumberToWordsPersian(ANumber div Hezar) + ' هزار و '
          + NumberToWordsPersian(ANumber mod Hezar);
    end
  else if (ANumber >= 1 * Milion) and (ANumber < 1 * Miliard) then
    begin
      if ANumber mod Milion = 0 then
        Result := NumberToWordsPersian(ANumber div Milion) + ' میلیون'
      else
        Result := NumberToWordsPersian(ANumber div Milion) + ' میلیون و '
          + NumberToWordsPersian(ANumber mod Milion);
    end
  else if (ANumber >= 1 * Miliard)then
    begin
      if ANumber mod Miliard = 0 then
        Result := NumberToWordsPersian(ANumber div Miliard) + ' میلیارد'
      else
        Result := NumberToWordsPersian(ANumber div Miliard) + ' میلیارد و '
          + NumberToWordsPersian(ANumber mod Miliard);
    end;
end;

class function TNumberUtility.NumberToWordsEnglish(ANumber: Int64): string;
const
  Hundred = 100;
  Thousand = 1000;
  Million = Thousand * 1000;
  Billion = Million * 1000;
  Tens: array[2..9] of string = ('twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety');
  Teens: array [10..19] of string = ('ten','eleven','twelve','thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen');
  Units: array [1..9] of string = ('one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine');
begin
  if (ANumber < 0) then Result := 'minus ' else Result := '';
  ANumber := abs(ANumber);  //Get ANumber as Positive Value;

  //Billions
  if (ANumber >= Billion) then
  begin
    Result := Result + NumberToWordsEnglish(ANumber div Billion) + ' bllion, ';
    ANumber := ANumber mod Billion;
  end;

  //Millions
  if (ANumber >= Million) then
  begin
    Result := Result + NumberToWordsEnglish(ANumber div Million) + ' million, ';
    ANumber := ANumber mod Million;
  end;

  //Thousands
  if (ANumber >= Thousand) then
  begin
    Result := Result + NumberToWordsEnglish(ANumber div Thousand) + ' thousand, ';
    ANumber := ANumber mod Thousand;
  end;

  //Hundreds
  if (ANumber >= Hundred) then
  begin
    Result := Result + NumberToWordsEnglish(ANumber div Hundred) + ' hundred';
    ANumber := ANumber mod Hundred;
  end;

  // and ...
  if (ANumber > 0)and(Result <> '') then
  begin
    Result := Trim(Result);
    if (Result[Length(Result)] = ',') then Delete(Result, Length(Result), 1);
    Result := Result + ' and ';
  end;

  //Tens
  if (ANumber >= 20) then
  begin
    Result := Result + Tens[ANumber div 10] + ' ';
    ANumber := ANumber mod 10;
  end;
  if (ANumber >= 10) then
  begin
    Result := Result + Teens[ANumber];
    ANumber := 0
  end;

  //Units
  if (ANumber >= 1) then Result := Result + Units[ANumber];

  //Tidy up the result
  Result := Trim(Result);
  if (Result = '') then Result := 'zero'
  else
    if (Result[Length(Result)] = ',') then
      Delete(Result, Length(Result), 1);
end;

{ TAlign }

class procedure TAlign.AlignControl(AControl: TControl; AlignType: SetOfAlignType);
var
  // اندازه ی والد شامل عرض و ارتفاع
  // Size of parent control include width and height
  ParentSize,
  // مختصات والد شامل سمت چپ و بالا
  // Position of parent control include left and top
  ParentPosition:TPoint;
begin
  if not Assigned(AControl.Parent) then
    Exit;

  ParentSize.x := AControl.Parent.ClientWidth;
  ParentSize.y := AControl.Parent.ClientHeight;
  ParentPosition.x:=AControl.Parent.Left;
  ParentPosition.y:=AControl.Parent.Top;

  if altLeft in AlignType then
  begin
    AControl.Left:=0;
    AControl.Top:=0;
    AControl.Height:=ParentSize.y;
  end;
  if altTop in AlignType then
  begin
    AControl.Left := 0;
    AControl.Top:=0;
    AControl.Width:=ParentSize.x;
  end;
  if altRight in AlignType then
  begin
    AControl.Left:=ParentSize.x - AControl.Width;
    AControl.Top:=0;
    AControl.Height:=ParentSize.y;
  end;
  if   altBottom in AlignType then
  begin
    AControl.Left := 0;
    AControl.Top:=ParentSize.y - AControl.Width;
    AControl.Width:=ParentSize.x;
  end;
  if altVerticallyCenter in AlignType then
  begin
    AControl.Top:=(ParentSize.y - AControl.Height) div 2;
  end;
  if altHorizontallyCenter in AlignType then
  begin
    AControl.Left:=(ParentSize.x - AControl.Width) div 2;
  end;
  if altCenter in AlignType then
  begin
    AControl.Left:=(ParentSize.x - AControl.Width) div 2;
    AControl.Top:=(ParentSize.y - AControl.Height) div 2;
  end;
end;

class procedure TAlign.AlignControls(AControlsContainer: TControl;
  Controls: array of TControl; AlignType: SetOfAlignType);
var
  SumOfControlsWidth, SumOfControlsHeight: integer;
  i, j, Space, CurrentPosition: integer;
  ControlsFromTop, ControlsFromLeft: array of TControl;
  TempControl: TControl;
begin
  // اگر چیدمان معمولی بود به تابع خود فرستاده می شود
  // If allign
  if (altTop in AlignType) or (altBottom in AlignType)
     or (altLeft in AlignType) or (altRight in AlignType)
     or (altCenter in AlignType) or (altHorizontallyCenter in AlignType)
     or (altVerticallyCenter in AlignType) then
    for i := low(Controls) to High(Controls) do
      AlignControl(Controls[i], AlignType);
  // Initial sum to zero
  SumOfControlsWidth:= 0;
  // Initial sum to zero
  SumOfControlsHeight:= 0;
  // یافتن بلندا و پهنای همه ی کنترل ها
  // Add widths and heights of any controls and calculate sum of them
  for i := Low(Controls) to High(Controls) do
  begin
    if Controls[i].Parent = AControlsContainer then
    begin
      SumOfControlsWidth:=  SumOfControlsWidth + Controls[i].Width;
      SumOfControlsHeight:= SumOfControlsHeight+ Controls[i].Height;
    end;
  end;
  // مرتب سازی همه ی کنترل ها بر پایه ی سمت چپ
  // Sort controlls with lefts properties
  for i := Low(Controls) to High(Controls) do
    for j := i + 1 to High(Controls) do
      if Controls[i].Parent = AControlsContainer then
        if Controls[i].Left > Controls[j].Left then
        begin
          TempControl := Controls[i];
          Controls[i] := Controls[j];
          Controls[j] := TempControl;
        end;
  // چیدمان فاصله ی برابر در راستای افق
  // Horizontally space equally between controls
  if altHorizontallySpaceEqually in AlignType then
  begin
    // یافتن اندازه ی فاصله ی میانگین برای کنترل ها
    // Calculate average space between scontrols
    Space:=(AControlsContainer.Width - SumOfControlsWidth)
            div (Length(Controls) + 1);
    CurrentPosition:=Space;
    for i := Low(Controls) to High(Controls) do
    begin
      if Controls[i].Parent = AControlsContainer then
      begin
        Controls[i].Left:= CurrentPosition;
        CurrentPosition := Controls[i].Width + CurrentPosition + Space;
      end;
    end;
  end;
  // مرتب سازی همه ی کنترل ها بر پایه ی سمت بالا
  // Sort controlls with tops properties
  for i := Low(Controls) to High(Controls) do
    for j := i + 1 to High(Controls) do
      if Controls[i].Parent = AControlsContainer then
        if Controls[i].Top > Controls[j].Top then
        begin
          TempControl := Controls[i];
          Controls[i] := Controls[j];
          Controls[j] := TempControl;
        end;
  // چیدمان فاصله ی برابر در راستای عمود
  // Vertically space equally between controls
  if altVerticallySpaceEqually in AlignType then
  begin
    // یافتن اندازه ی فاصله ی میانگین برای کنترل ها
    // Calculate average space between controls
    Space:=(AControlsContainer.Height - SumOfControlsHeight)
            div (Length(Controls) + 1);
    CurrentPosition:=Space;
    for i := Low(Controls) to High(Controls) do
    begin
      if Controls[i].Parent = AControlsContainer then
      begin
        Controls[i].Top:= CurrentPosition;
        CurrentPosition := Controls[i].Height + CurrentPosition + Space;
      end;
    end;
  end;
  //for
end;

class procedure TAlign.AlignControls(AControlsContainer: TWinControl;
  AlignType: SetOfAlignType);
var
  Controls: array of TControl;
  i: integer;
begin
  SetLength(Controls, AControlsContainer.ControlCount);
  for i := 0 to AControlsContainer.ControlCount - 1 do
    Controls[i] := AControlsContainer.Controls[i];
  AlignControls(AControlsContainer, Controls, AlignType);
end;

{ TSizing }

class procedure TSizing.ResizeControl(AControl: TControl; SizeType: TSizeType);
begin
    if not Assigned(AControl.Parent) then
    Exit;

  case SizeType of
    stParentWidth:
      begin
        AControl.Width := AControl.Parent.ClientWidth;
      end;
    stParentHeight:
      begin
        AControl.Height:=AControl.Parent.ClientHeight;
      end;
    stParentSize:
      begin
        AControl.Width:=AControl.Parent.ClientWidth;
        AControl.Height:=AControl.Parent.ClientHeight;
      end;
    stParentHalfWidth:
      begin
        AControl.Width := AControl.Parent.ClientWidth div 2;
      end;
    stParentHalfHeight:
      begin
        AControl.Height:=AControl.Parent.ClientHeight div 2;
      end;
    stParentHalfsize:
      begin
        AControl.Width := AControl.Parent.ClientWidth div 2;
        AControl.Height:=AControl.Parent.ClientHeight div 2;
      end;
    stParentOneThirdWidth:
      begin
        AControl.Width := AControl.Parent.ClientWidth div 3;
      end;
    stParentOneThirdHeight:
      begin
        AControl.Height:=AControl.Parent.ClientHeight div 3;
      end;
    stParentOneThirdSize:
      begin
        AControl.Width := AControl.Parent.ClientWidth div 3;
        AControl.Height:=AControl.Parent.ClientHeight div 3;
      end;
    stParentQuarterWidth:
      begin
        AControl.Width := AControl.Parent.ClientWidth div 4;
      end;
    stParentQuarterHeight:
      begin
        AControl.Height:=AControl.Parent.ClientHeight div 4;
      end;
    stParentQuarterSize:
      begin
        AControl.Width := AControl.Parent.ClientWidth div 4;
        AControl.Height:=AControl.Parent.ClientHeight div 4;
      end;
  end;
end;

class procedure TSizing.ResizeControl(AControl: TControl; AWidth,
  AHeight: Integer);
begin
  AControl.Width:=AWidth;
  AControl.Height:=AHeight;
end;

class procedure TSizing.ResizeGridColumnsWidth(AGrid: TDrawGrid;
  ColumnsWidth: Word);
begin
  AGrid.DefaultColWidth := ColumnsWidth;
end;

class procedure TSizing.ResizeGridColumnsWidth(AGrid: TDrawGrid;
  ColumnsWidth: array of Word);
var
  I: integer;
begin
  for I := 0 to AGrid.ColCount - 1 do
    if Length(ColumnsWidth) > i then
      AGrid.ColWidths[i] := ColumnsWidth[i];
end;

class procedure TSizing.ResizeGridRowsHeight(AGrid: TDrawGrid;
  RowsHeight: Word);
begin
  AGrid.DefaultColWidth := RowsHeight;
end;

class procedure TSizing.ResizeGridRowsHeight(AGrid: TDrawGrid;
  RowsHeight: array of Word);
var
  I: Integer;
begin
  for I := 0 to AGrid.RowCount - 1 do
    if Length(RowsHeight) > i then
      AGrid.RowHeights[i] := RowsHeight[i];
end;

{ TStringHelper }
{$IfNDef FPC}
function TStringHelper.GetNumber: string;
begin
  Result := StringReplace(Self, FormatSettings.ThousandSeparator, ''
    , [rfReplaceAll]);
end;


{ TEditHelper }

function TEditHelper.GetNumber: string;
begin
  Result := string(Self.Text).Number;
end;
{$EndIf}
{ TFontUtility }

{$IFDEF MSWINDOWS}
function EnumFontsProc(var LogFont: TLogFont; var TextMetric: TTextMetric;
  FontType: Integer; Data: Pointer): Integer; stdcall;
var
  S: TStrings;
  Temp: string;
begin
  S := TStrings(Data);
  Temp := LogFont.lfFaceName;
  if (S.Count = 0) or (AnsiCompareText(S[S.Count-1], Temp) <> 0) then
    S.Add(Temp);
  Result := 1;
end;
{$ENDIF}

{$IFNDEF FPC}
class function TFontUtility.CollectFonts: TStrings;
var
{$IFDEF MACOS}
  fManager: NsFontManager;
  list:NSArray;
  lItem:NSString;
{$ENDIF}
{$IFDEF MSWINDOWS}
  DC: HDC;
  LFont: TLogFont;
{$ENDIF}
  i: Integer;
begin
  Result := TStringList.Create;

  {$IFDEF MACOS}
    fManager := TNsFontManager.Wrap(TNsFontManager.OCClass.sharedFontManager);
    list := fManager.availableFontFamilies;
    if (List <> nil) and (List.count > 0) then
    begin
      for i := 0 to List.Count-1 do
      begin
        lItem := TNSString.Wrap(List.objectAtIndex(i));
        Result.Add(String(lItem.UTF8String))
      end;
    end;
  {$ENDIF}
  {$IFDEF MSWINDOWS}
    DC := GetDC(0);
    FillChar(LFont, sizeof(LFont), 0);
    LFont.lfCharset := DEFAULT_CHARSET;
    EnumFontFamiliesEx(DC, LFont, @EnumFontsProc, Winapi.Windows.LPARAM(Result), 0);
    ReleaseDC(0, DC);
  {$ENDIF}
end;
{$ENDIF}

end.
