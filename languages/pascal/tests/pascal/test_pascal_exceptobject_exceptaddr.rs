use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 56: RTL Exception Inspection (ExceptObject & ExceptAddr)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_exceptobject_basic_query() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    raise Exception.Create('ExceptObjectTest');
  except
    WriteLn(Exception(ExceptObject).Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["ExceptObjectTest"]);
}

#[test]
fn test_exceptobject_classname() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    raise EInvalidArgument.Create('BadArg');
  except
    WriteLn(ExceptObject.ClassName);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["EInvalidArgument"]);
}

#[test]
fn test_exceptaddr_not_nil() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    raise Exception.Create('CheckAddr');
  except
    WriteLn(ExceptAddr <> nil);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_exceptobject_casting_to_custom_exception() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ECustom = class(Exception)
  public Tag: Integer;
  constructor CreateTag(T: Integer; const msg: String);
end;
constructor ECustom.CreateTag(T: Integer; const msg: String);
begin
  inherited Create(msg); Tag := T;
end;
begin
  try
    raise ECustom.CreateTag(777, 'CustomTagged');
  except
    if ExceptObject is ECustom then
      WriteLn(ECustom(ExceptObject).Tag);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["777"]);
}

#[test]
fn test_exceptobject_outside_except_is_nil() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  WriteLn(ExceptObject = nil);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_pass_exceptobject_to_logger() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure LogException(obj: TObject);
begin
  if obj is Exception then
    WriteLn('Logged:' + Exception(obj).Message);
end;
begin
  try
    raise Exception.Create('PassToLogger');
  except
    LogException(ExceptObject);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["Logged:PassToLogger"]);
}

#[test]
fn test_pass_exceptaddr_to_reporter() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure ReportAddr(addr: Pointer);
begin
  WriteLn(addr <> nil);
end;
begin
  try
    raise Exception.Create('PassAddr');
  except
    ReportAddr(ExceptAddr);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_exceptobject_in_untyped_except_block() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    raise EDivByZero.Create('UntypedDivZero');
  except
    WriteLn(ExceptObject.ClassName + ':' + Exception(ExceptObject).Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["EDivByZero:UntypedDivZero"]);
}

#[test]
fn test_exceptobject_nested_handlers() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    try
      raise Exception.Create('InnerExc');
    except
      WriteLn('InnerExceptObject:' + Exception(ExceptObject).Message);
      raise Exception.Create('OuterExc');
    end;
  except
    WriteLn('OuterExceptObject:' + Exception(ExceptObject).Message);
  end;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["InnerExceptObject:InnerExc", "OuterExceptObject:OuterExc"]
    );
}

#[test]
fn test_exceptaddr_in_nested_handlers() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var a1, a2: Pointer;
begin
  try
    try
      raise Exception.Create('E1');
    except
      a1 := ExceptAddr;
      raise Exception.Create('E2');
    end;
  except
    a2 := ExceptAddr;
  end;
  WriteLn(a1 <> nil);
  WriteLn(a2 <> nil);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_exceptobject_in_constructor_failure() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TFail = class
  constructor Create;
end;
constructor TFail.Create;
begin
  raise Exception.Create('CtorErr');
end;
var obj: TFail;
begin
  try
    obj := TFail.Create;
  except
    WriteLn(ExceptObject.ClassName);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["Exception"]);
}

#[test]
fn test_exceptobject_in_record_method() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type TRec = record
  procedure Fail;
end;
procedure TRec.Fail;
begin
  raise Exception.Create('RecMethodFail');
end;
var r: TRec;
begin
  try
    r.Fail;
  except
    WriteLn(Exception(ExceptObject).Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["RecMethodFail"]);
}

#[test]
fn test_exceptobject_inheritsfrom_check() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
type ECustomBase = class(Exception);
type ECustomSub = class(ECustomBase);
begin
  try
    raise ECustomSub.Create('SubClassError');
  except
    WriteLn(ExceptObject.InheritsFrom(ECustomBase));
  end;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_exceptaddr_converted_to_hex_string() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var addrHex: String;
begin
  try
    raise Exception.Create('AddrHexTest');
  except
    addrHex := HexStr(NativeInt(ExceptAddr), 8);
    WriteLn(Length(addrHex) >= 8);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_exceptobject_property_access() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    raise Exception.Create('PropAccessMessage');
  except
    WriteLn(Exception(ExceptObject).Message);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["PropAccessMessage"]);
}

#[test]
fn test_exceptobject_reraised_identity() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var ptr1, ptr2: Pointer;
begin
  try
    try
      raise Exception.Create('IdentityTest');
    except
      ptr1 := Pointer(ExceptObject);
      raise;
    end;
  except
    ptr2 := Pointer(ExceptObject);
  end;
  WriteLn(ptr1 = ptr2);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_exceptobject_in_function_return() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
function GetExceptMsg: String;
begin
  try
    raise Exception.Create('FuncExceptMsg');
  except
    Result := Exception(ExceptObject).Message;
  end;
end;
begin
  WriteLn(GetExceptMsg);
end.
"#,
    );
    assert_eq!(out, vec!["FuncExceptMsg"]);
}

#[test]
fn test_exceptobject_with_sysutils_econvert_error() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
begin
  try
    StrToInt('Invalid');
  except
    WriteLn(ExceptObject.ClassName);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["EConvertError"]);
}

#[test]
fn test_exceptobject_with_edivbyzero() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
var a, b: Integer;
begin
  a := 5; b := 0;
  try
    a := a div b;
  except
    WriteLn(ExceptObject.ClassName);
  end;
end.
"#,
    );
    assert_eq!(out, vec!["EDivByZero"]);
}

#[test]
fn test_exceptobject_lifecycle_safety() {
    let out = run_pascal(
        r#"
program Test;
uses SysUtils;
procedure InspectObject;
begin
  WriteLn(ExceptObject <> nil);
end;
begin
  try
    raise Exception.Create('LifecycleSafety');
  except
    InspectObject;
  end;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}
