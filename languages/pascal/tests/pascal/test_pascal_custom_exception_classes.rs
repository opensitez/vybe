use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 54: Custom Exception Classes & Inheritance Hierarchies
// ═══════════════════════════════════════════════════════════

#[test]
fn test_custom_exception_class_declaration() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ECustomAppError = class(Exception);
begin
  try
    raise ECustomAppError.Create('AppErrorOccurred');
  except
    on E: ECustomAppError do WriteLn(E.ClassName + ':' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["ECustomAppError:AppErrorOccurred"]);
}

#[test]
fn test_custom_exception_with_error_code_field() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ECodeException = class(Exception)
  private FErrorCode: Integer;
  public constructor CreateCode(code: Integer; const msg: String);
  public property ErrorCode: Integer read FErrorCode;
end;
constructor ECodeException.CreateCode(code: Integer; const msg: String);
begin
  inherited Create(msg);
  FErrorCode := code;
end;
begin
  try
    raise ECodeException.CreateCode(404, 'Page Not Found');
  except
    on E: ECodeException do WriteLn(E.ErrorCode.ToString + '-' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["404-Page Not Found"]);
}

#[test]
fn test_custom_exception_inheritance_polymorphism() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type EBaseAppErr = class(Exception);
type ENetErr = class(EBaseAppErr);
type ETimeoutErr = class(ENetErr);

procedure TriggerTimeout;
begin
  raise ETimeoutErr.Create('ConnectionTimedOut');
end;
begin
  try
    TriggerTimeout;
  except
    on E: EBaseAppErr do WriteLn('CaughtBase:' + E.ClassName);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["CaughtBase:ETimeoutErr"]);
}

#[test]
fn test_custom_exception_subclass_matching_priority() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type EBaseErr = class(Exception);
type ESubErr = class(EBaseErr);

begin
  try
    raise ESubErr.Create('SubError');
  except
    on E: ESubErr do WriteLn('MatchedSubClass');
    on E: EBaseErr do WriteLn('MatchedBaseClass');
  end;
end.
"#,
    );
    assert_eq!(out, vec!["MatchedSubClass"]);
}

#[test]
fn test_custom_exception_with_default_constructor_args() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type EDefaultErr = class(Exception)
  public constructor CreateDefault(msg: String = 'DefaultErrorMessage');
end;
constructor EDefaultErr.CreateDefault(msg: String);
begin
  inherited Create(msg);
end;
begin
  try
    raise EDefaultErr.CreateDefault;
  except
    on E: EDefaultErr do WriteLn(E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["DefaultErrorMessage"]);
}

#[test]
fn test_custom_exception_inheritsfrom_query() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ECoreErr = class(Exception);
type EValErr = class(ECoreErr);
var e: Exception;
begin
  e := EValErr.Create('ValidationError');
  WriteLn(e.InheritsFrom(ECoreErr));
  WriteLn(e.InheritsFrom(Exception));
  e.Free;
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_custom_exception_thrown_from_property_setter() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type EInvalidAge = class(Exception);
type TPerson = class
  private FAge: Integer;
  private procedure SetAge(v: Integer);
  public property Age: Integer read FAge write SetAge;
end;
procedure TPerson.SetAge(v: Integer);
begin
  if v < 0 then raise EInvalidAge.Create('AgeCannotBeNegative');
  FAge := v;
end;
var p: TPerson;
begin
  p := TPerson.Create;
  try
    p.Age := -10;
  except
    on E: EInvalidAge do WriteLn(E.Message);
  end;
  p.Free;
end.
"#,
    );
    assert_eq!(out, vec!["AgeCannotBeNegative"]);
}

#[test]
fn test_custom_exception_createfmt_helper() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type EFormattedErr = class(Exception);
begin
  try
    raise EFormattedErr.CreateFmt('Invalid value %d for field %s', [99, 'Age']);
  except
    on E: EFormattedErr do WriteLn(E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["Invalid value 99 for field Age"]);
}

#[test]
fn test_custom_exception_with_helpcontext() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type EHelpErr = class(Exception)
  public constructor CreateHelp(const msg: String; helpCtx: Integer);
end;
constructor EHelpErr.CreateHelp(const msg: String; helpCtx: Integer);
begin
  inherited CreateHelp(msg, helpCtx);
end;
begin
  try
    raise EHelpErr.CreateHelp('ContextualHelpError', 1001);
  except
    on E: EHelpErr do WriteLn(E.Message + '-' + E.HelpContext.ToString);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["ContextualHelpError-1001"]);
}

#[test]
fn test_custom_exception_wrapping_payload_record() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TErrorPayload = record ModuleName: String; Severity: Integer; end;
type EPayloadErr = class(Exception)
  public Payload: TErrorPayload;
  constructor CreatePayload(P: TErrorPayload);
end;
constructor EPayloadErr.CreatePayload(P: TErrorPayload);
begin
  inherited Create('PayloadError');
  Payload := P;
end;
var pData: TErrorPayload;
begin
  pData.ModuleName := 'AuthModule'; pData.Severity := 5;
  try
    raise EPayloadErr.CreatePayload(pData);
  except
    on E: EPayloadErr do WriteLn(E.Payload.ModuleName + ':' + E.Payload.Severity.ToString);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["AuthModule:5"]);
}

#[test]
fn test_custom_exception_virtual_message_formatter() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type EVirtualErr = class(Exception)
  public function GetFormattedMessage: String; virtual;
end;
type ESubVirtualErr = class(EVirtualErr)
  public function GetFormattedMessage: String; override;
end;
function EVirtualErr.GetFormattedMessage: String; begin Result := 'Base:' + Message; end;
function ESubVirtualErr.GetFormattedMessage: String; begin Result := 'Sub:' + Message; end;

var e: EVirtualErr;
begin
  e := ESubVirtualErr.Create('TestMessage');
  WriteLn(e.GetFormattedMessage);
  e.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Sub:TestMessage"]);
}

#[test]
fn test_custom_exception_in_generic_method() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type EGenericErr = class(Exception);
type TValidator = class
  public class procedure CheckNotNull<T>(const val: T);
end;
class procedure TValidator.CheckNotNull<T>(const val: T);
begin
  // Mock check
  raise EGenericErr.Create('GenericValueNull');
end;
begin
  try
    TValidator.CheckNotNull<String>('');
  except
    on E: EGenericErr do WriteLn(E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["GenericValueNull"]);
}

#[test]
fn test_custom_exception_classname_reflection() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ECustomReflect = class(Exception);
var e: ECustomReflect;
begin
  e := ECustomReflect.Create('Reflect');
  WriteLn(e.ClassName);
  e.Free;
end.
"#,
    );
    assert_eq!(out, vec!["ECustomReflect"]);
}

#[test]
fn test_custom_exception_in_record_method() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ERecordErr = class(Exception);
type TRec = record
  procedure Validate;
end;
procedure TRec.Validate;
begin
  raise ERecordErr.Create('RecordValidationFailed');
end;
var r: TRec;
begin
  try
    r.Validate;
  except
    on E: ERecordErr do WriteLn(E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["RecordValidationFailed"]);
}

#[test]
fn test_custom_exception_hierarchy_multi_level() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ELevel1 = class(Exception);
type ELevel2 = class(ELevel1);
type ELevel3 = class(ELevel2);

procedure RunLevel3;
begin
  raise ELevel3.Create('L3');
end;

begin
  try
    RunLevel3;
  except
    on E: ELevel2 do WriteLn('CaughtAtL2:' + E.ClassName);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["CaughtAtL2:ELevel3"]);
}

#[test]
fn test_custom_exception_multiple_fields_constructor() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type EMultiFieldErr = class(Exception)
  public Line, Col: Integer;
  constructor CreatePos(L, C: Integer; const msg: String);
end;
constructor EMultiFieldErr.CreatePos(L, C: Integer; const msg: String);
begin
  inherited Create(msg);
  Line := L; Col := C;
end;
begin
  try
    raise EMultiFieldErr.CreatePos(12, 45, 'SyntaxError');
  except
    on E: EMultiFieldErr do WriteLn(E.Message + '@' + E.Line.ToString + ':' + E.Col.ToString);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["SyntaxError@12:45"]);
}

#[test]
fn test_custom_exception_subclassing_econvert_error() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type EParseErr = class(EConvertError);
begin
  try
    raise EParseErr.Create('ParseIntegerFailed');
  except
    on E: EConvertError do WriteLn('CaughtAsConvertError:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["CaughtAsConvertError:ParseIntegerFailed"]);
}

#[test]
fn test_custom_exception_reraising_custom_type() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ECustomReraise = class(Exception);
procedure Inner;
begin
  raise ECustomReraise.Create('ReraisedCustom');
end;
begin
  try
    try
      Inner;
    except
      on E: ECustomReraise do
      begin
        WriteLn('InnerLog:' + E.Message);
        raise;
      end;
    end;
  except
    on E: ECustomReraise do WriteLn('OuterLog:' + E.Message);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["InnerLog:ReraisedCustom", "OuterLog:ReraisedCustom"]
    );
}

#[test]
fn test_custom_exception_array_instantiation() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type EArrayErr = class(Exception);
var errs: array[0..1] of EArrayErr;
begin
  errs[0] := EArrayErr.Create('Err0');
  errs[1] := EArrayErr.Create('Err1');
  WriteLn(errs[0].Message);
  WriteLn(errs[1].Message);
  errs[0].Free; errs[1].Free;
end.
"#,
    );
    assert_eq!(out, vec!["Err0", "Err1"]);
}

#[test]
fn test_custom_exception_with_string_property() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type EDetailErr = class(Exception)
  private FDetail: String;
  public constructor CreateDetail(const msg, detail: String);
  public property Detail: String read FDetail;
end;
constructor EDetailErr.CreateDetail(const msg, detail: String);
begin
  inherited Create(msg);
  FDetail := detail;
end;
begin
  try
    raise EDetailErr.CreateDetail('Summary', 'DetailedStackInformation');
  except
    on E: EDetailErr do WriteLn(E.Message + ' -> ' + E.Detail);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["Summary -> DetailedStackInformation"]);
}
