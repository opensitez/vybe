use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 58: Exception Class Polymorphic Matching (on E: Class do)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_exact_class_matching() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  try
    raise EConvertError.Create('InvalidInt');
  except
    on E: EConvertError do WriteLn('MatchedExact:' + E.ClassName);
    on E: Exception do WriteLn('MatchedBase');
  end;
end.
"#);
    assert_eq!(out, vec!["MatchedExact:EConvertError"]);
}

#[test]
fn test_subclass_matches_ancestor_handler() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type ECustomBase = class(Exception);
type ECustomChild = class(ECustomBase);
begin
  try
    raise ECustomChild.Create('ChildFail');
  except
    on E: ECustomBase do WriteLn('MatchedAncestor:' + E.ClassName);
  end;
end.
"#);
    assert_eq!(out, vec!["MatchedAncestor:ECustomChild"]);
}

#[test]
fn test_first_matching_on_block_wins() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type EBase = class(Exception);
type ESub = class(EBase);
begin
  try
    raise ESub.Create('SubError');
  except
    on E: EBase do WriteLn('BaseHandlerFirst');
    on E: ESub do WriteLn('SubHandlerSecond');
  end;
end.
"#);
    assert_eq!(out, vec!["BaseHandlerFirst"]);
}

#[test]
fn test_omitting_variable_in_on_clause() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  try
    raise EDivByZero.Create('DivideByZero');
  except
    on EDivByZero do WriteLn('DivByZeroCaughtWithoutVar');
  end;
end.
"#);
    assert_eq!(out, vec!["DivByZeroCaughtWithoutVar"]);
}

#[test]
fn test_on_block_with_else_fallback() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  try
    raise ERangeError.Create('OutBits');
  except
    on E: EDivByZero do WriteLn('DivZero');
    on E: EConvertError do WriteLn('ConvertErr');
  else
    WriteLn('ElseFallbackHandled');
  end;
end.
"#);
    assert_eq!(out, vec!["ElseFallbackHandled"]);
}

#[test]
fn test_multi_level_inheritance_matching() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type ELevel1 = class(Exception);
type ELevel2 = class(ELevel1);
type ELevel3 = class(ELevel2);
begin
  try
    raise ELevel3.Create('L3');
  except
    on E: ELevel1 do WriteLn('CaughtAtL1:' + E.ClassName);
  end;
end.
"#);
    assert_eq!(out, vec!["CaughtAtL1:ELevel3"]);
}

#[test]
fn test_on_exception_class_eaccessviolation() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var p: PInteger;
begin
  p := nil;
  try
    p^ := 10;
  except
    on E: EAccessViolation do WriteLn('MatchedAV');
    on E: Exception do WriteLn('MatchedGen');
  end;
end.
"#);
    assert_eq!(out, vec!["MatchedAV"]);
}

#[test]
fn test_on_exception_class_edivbyzero() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var a, b: Integer;
begin
  a := 10; b := 0;
  try
    a := a div b;
  except
    on E: EDivByZero do WriteLn('MatchedDivByZero');
  end;
end.
"#);
    assert_eq!(out, vec!["MatchedDivByZero"]);
}

#[test]
fn test_on_exception_class_eargumentexception() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  try
    raise EArgumentException.Create('BadParam');
  except
    on E: EArgumentException do WriteLn('MatchedArgErr:' + E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["MatchedArgErr:BadParam"]);
}

#[test]
fn test_on_exception_custom_property_access() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type ECustomProp = class(Exception)
  public CustomID: Integer;
  constructor CreateID(id: Integer; msg: String);
end;
constructor ECustomProp.CreateID(id: Integer; msg: String);
begin
  inherited Create(msg); CustomID := id;
end;
begin
  try
    raise ECustomProp.CreateID(888, 'PropMsg');
  except
    on E: ECustomProp do WriteLn(E.CustomID.ToString + '-' + E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["888-PropMsg"]);
}

#[test]
fn test_on_exception_in_virtual_method_override() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type TBaseWorker = class
  public procedure DoWork; virtual;
end;
type TSubWorker = class(TBaseWorker)
  public procedure DoWork; override;
end;
procedure TBaseWorker.DoWork; begin end;
procedure TSubWorker.DoWork; begin raise EInvalidOp.Create('SubOpFail'); end;

var w: TBaseWorker;
begin
  w := TSubWorker.Create;
  try
    w.DoWork;
  except
    on E: EInvalidOp do WriteLn('MatchedVirtualOverride:' + E.Message);
  end;
  w.Free;
end.
"#);
    assert_eq!(out, vec!["MatchedVirtualOverride:SubOpFail"]);
}

#[test]
fn test_on_exception_in_interface_delegation() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type IWork = interface
  ['{11111111-2222-3333-4444-555555555555}']
  procedure Perform;
end;
type TWorkImpl = class(TInterfacedObject, IWork)
  public procedure Perform;
end;
procedure TWorkImpl.Perform; begin raise EOverflow.Create('IntOverflow'); end;

var w: IWork;
begin
  w := TWorkImpl.Create;
  try
    w.Perform;
  except
    on E: EOverflow do WriteLn('MatchedInterfaceDelegation:' + E.ClassName);
  end;
end.
"#);
    assert_eq!(out, vec!["MatchedInterfaceDelegation:EOverflow"]);
}

#[test]
fn test_on_exception_in_constructor() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type TFailObj = class
  constructor Create;
end;
constructor TFailObj.Create; begin raise ERangeError.Create('RangeCtor'); end;

var obj: TFailObj;
begin
  try
    obj := TFailObj.Create;
  except
    on E: ERangeError do WriteLn('MatchedCtorErr:' + E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["MatchedCtorErr:RangeCtor"]);
}

#[test]
fn test_on_exception_in_record_method() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type TRec = record
  procedure Run;
end;
procedure TRec.Run; begin raise EConvertError.Create('RecConvertFail'); end;

var r: TRec;
begin
  try
    r.Run;
  except
    on E: EConvertError do WriteLn('MatchedRecordErr:' + E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["MatchedRecordErr:RecConvertFail"]);
}

#[test]
fn test_on_exception_variable_scope_isolation() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure TestScope;
begin
  try
    raise Exception.Create('Err1');
  except
    on E: Exception do WriteLn('Scope1:' + E.Message);
  end;

  try
    raise Exception.Create('Err2');
  except
    on E: Exception do WriteLn('Scope2:' + E.Message);
  end;
end;
begin
  TestScope;
end.
"#);
    assert_eq!(out, vec!["Scope1:Err1", "Scope2:Err2"]);
}

#[test]
fn test_on_exception_reraise_from_handler() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise EDivByZero.Create('DivByZeroMessage');
    except
      on E: EDivByZero do
      begin
        WriteLn('InnerMatched:' + E.ClassName);
        raise;
      end;
    end;
  except
    on E: EDivByZero do WriteLn('OuterMatched:' + E.ClassName);
  end;
end.
"#);
    assert_eq!(out, vec!["InnerMatched:EDivByZero", "OuterMatched:EDivByZero"]);
}

#[test]
fn test_on_exception_catch_all_base_class() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type EUnusualError = class(Exception);
begin
  try
    raise EUnusualError.Create('Unusual');
  except
    on E: Exception do WriteLn('CatchAllBase:' + E.ClassName);
  end;
end.
"#);
    assert_eq!(out, vec!["CatchAllBase:EUnusualError"]);
}

#[test]
fn test_on_exception_ordered_specific_to_generic() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type ESpecific = class(Exception);
begin
  try
    raise ESpecific.Create('SpecificErr');
  except
    on E: ESpecific do WriteLn('SpecificHandlerMatched');
    on E: Exception do WriteLn('GenericHandlerMatched');
  end;
end.
"#);
    assert_eq!(out, vec!["SpecificHandlerMatched"]);
}

#[test]
fn test_on_exception_loop_matching_performance() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var i, count: Integer;
begin
  count := 0;
  for i := 1 to 3 do
  begin
    try
      raise EDivByZero.Create('DivZero');
    except
      on E: EDivByZero do Inc(count);
    end;
  end;
  WriteLn(count);
end.
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_on_exception_in_nested_function_calls() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
function F3: Integer; begin raise EConvertError.Create('F3Convert'); end;
function F2: Integer; begin Result := F3; end;
function F1: Integer; begin Result := F2; end;
begin
  try
    F1;
  except
    on E: EConvertError do WriteLn('TopMatchedF3:' + E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["TopMatchedF3:F3Convert"]);
}
