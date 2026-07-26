use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 16: Class Helpers & Record Helpers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_record_helper_for_integer_method() {
    let out = run_pascal(
        r#"
program Test;
type TIntHelper = record helper for Integer
  function IsEven: Boolean;
end;
function TIntHelper.IsEven: Boolean;
begin
  Result := (Self mod 2) = 0;
end;
var n: Integer;
begin
  n := 42;
  WriteLn(n.IsEven);
  n := 17;
  WriteLn(n.IsEven);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_record_helper_for_string_method() {
    let out = run_pascal(
        r#"
program Test;
type TStringHelper = record helper for String
  function DoubleStr: String;
end;
function TStringHelper.DoubleStr: String;
begin
  Result := Self + Self;
end;
var s: String;
begin
  s := 'ABC';
  WriteLn(s.DoubleStr);
end.
"#,
    );
    assert_eq!(out, vec!["ABCABC"]);
}

#[test]
fn test_record_helper_for_boolean_method() {
    let out = run_pascal(
        r#"
program Test;
type TBoolHelper = record helper for Boolean
  function ToInt: Integer;
end;
function TBoolHelper.ToInt: Integer;
begin
  if Self then Result := 1 else Result := 0;
end;
var b: Boolean;
begin
  b := True;
  WriteLn(b.ToInt);
  b := False;
  WriteLn(b.ToInt);
end.
"#,
    );
    assert_eq!(out, vec!["1", "0"]);
}

#[test]
fn test_class_helper_for_tobject() {
    let out = run_pascal(
        r#"
program Test;
type TObjectHelper = class helper for TObject
  function GetClassNameUpper: String;
end;
function TObjectHelper.GetClassNameUpper: String;
begin
  Result := UpperCase(Self.ClassName);
end;
var obj: TObject;
begin
  obj := TObject.Create;
  WriteLn(obj.GetClassNameUpper);
  obj.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TOBJECT"]);
}

#[test]
fn test_class_helper_for_custom_class() {
    let out = run_pascal(
        r#"
program Test;
type TPerson = class
  public Name: String;
  constructor Create(N: String);
end;
type TPersonHelper = class helper for TPerson
  function Greet: String;
end;
constructor TPerson.Create(N: String); begin Name := N; end;
function TPersonHelper.Greet: String; begin Result := 'Hello ' + Self.Name; end;
var p: TPerson;
begin
  p := TPerson.Create('Bob');
  WriteLn(p.Greet);
  p.Free;
end.
"#,
    );
    assert_eq!(out, vec!["Hello Bob"]);
}

#[test]
fn test_record_helper_property() {
    let out = run_pascal(
        r#"
program Test;
type TIntPropHelper = record helper for Integer
  private function GetIsPositive: Boolean;
  public property IsPositive: Boolean read GetIsPositive;
end;
function TIntPropHelper.GetIsPositive: Boolean;
begin
  Result := Self > 0;
end;
var val: Integer;
begin
  val := 15;
  WriteLn(val.IsPositive);
  val := -5;
  WriteLn(val.IsPositive);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_record_helper_enum_extension() {
    let out = run_pascal(
        r#"
program Test;
type TColor = (cRed, cGreen, cBlue);
type TColorHelper = record helper for TColor
  function ToHex: String;
end;
function TColorHelper.ToHex: String;
begin
  case Self of
    cRed: Result := '#FF0000';
    cGreen: Result := '#00FF00';
    cBlue: Result := '#0000FF';
  end;
end;
var col: TColor;
begin
  col := cGreen;
  WriteLn(col.ToHex);
end.
"#,
    );
    assert_eq!(out, vec!["#00FF00"]);
}

#[test]
fn test_record_helper_multiple_methods() {
    let out = run_pascal(
        r#"
program Test;
type TMathHelper = record helper for Integer
  function Squared: Integer;
  function Cubed: Integer;
end;
function TMathHelper.Squared: Integer; begin Result := Self * Self; end;
function TMathHelper.Cubed: Integer; begin Result := Self * Self * Self; end;
var x: Integer;
begin
  x := 3;
  WriteLn(x.Squared);
  WriteLn(x.Cubed);
end.
"#,
    );
    assert_eq!(out, vec!["9", "27"]);
}

#[test]
fn test_record_helper_var_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TMutateHelper = record helper for Integer
  procedure DoubleSelf;
end;
procedure TMutateHelper.DoubleSelf;
begin
  Self := Self * 2;
end;
var num: Integer;
begin
  num := 21;
  num.DoubleSelf;
  WriteLn(num);
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_record_helper_float_methods() {
    let out = run_pascal(
        r#"
program Test;
type TFloatHelper = record helper for Real
  function RoundToInt: Integer;
end;
function TFloatHelper.RoundToInt: Integer;
begin
  Result := Round(Self);
end;
var r: Real;
begin
  r := 14.8;
  WriteLn(r.RoundToInt);
end.
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_record_helper_chaining() {
    let out = run_pascal(
        r#"
program Test;
type TIntHelper = record helper for Integer
  function IncBy(n: Integer): Integer;
end;
function TIntHelper.IncBy(n: Integer): Integer; begin Result := Self + n; end;
var val: Integer;
begin
  val := 10;
  WriteLn(val.IncBy(5).IncBy(20));
end.
"#,
    );
    assert_eq!(out, vec!["35"]);
}

#[test]
fn test_class_helper_accesses_protected_field() {
    let out = run_pascal(
        r#"
program Test;
type TWidget = class
  protected FCode: Integer;
  public constructor Create(C: Integer);
end;
type TWidgetHelper = class helper for TWidget
  function GetCodeTimesTen: Integer;
end;
constructor TWidget.Create(C: Integer); begin FCode := C; end;
function TWidgetHelper.GetCodeTimesTen: Integer; begin Result := Self.FCode * 10; end;
var w: TWidget;
begin
  w := TWidget.Create(5);
  WriteLn(w.GetCodeTimesTen);
  w.Free;
end.
"#,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_record_helper_default_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TStrPadHelper = record helper for String
  function PadLeftCustom(len: Integer; ch: Char = ' '): String;
end;
function TStrPadHelper.PadLeftCustom(len: Integer; ch: Char): String;
begin
  Result := Self;
  while Length(Result) < len do
    Result := ch + Result;
end;
var s: String;
begin
  s := '7';
  WriteLn(s.PadLeftCustom(3, '0'));
end.
"#,
    );
    assert_eq!(out, vec!["007"]);
}

#[test]
fn test_record_helper_overloaded_methods() {
    let out = run_pascal(
        r#"
program Test;
type TIntOverloadHelper = record helper for Integer
  function AddVal(n: Integer): Integer; overload;
  function AddVal(s: String): Integer; overload;
end;
function TIntOverloadHelper.AddVal(n: Integer): Integer; begin Result := Self + n; end;
function TIntOverloadHelper.AddVal(s: String): Integer; begin Result := Self + StrToInt(s); end;
var x: Integer;
begin
  x := 10;
  WriteLn(x.AddVal(5));
  WriteLn(x.AddVal('20'));
end.
"#,
    );
    assert_eq!(out, vec!["15", "30"]);
}

#[test]
fn test_record_helper_subrange_type() {
    let out = run_pascal(
        r#"
program Test;
type TScore = 0..100;
type TScoreHelper = record helper for TScore
  function IsPassing: Boolean;
end;
function TScoreHelper.IsPassing: Boolean; begin Result := Self >= 60; end;
var sc: TScore;
begin
  sc := 75;
  WriteLn(sc.IsPassing);
  sc := 45;
  WriteLn(sc.IsPassing);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_class_helper_static_class_method() {
    let out = run_pascal(
        r#"
program Test;
type TBase = class end;
type TBaseStaticHelper = class helper for TBase
  public class procedure Announce;
end;
class procedure TBaseStaticHelper.Announce;
begin
  WriteLn('HelperStaticMethod');
end;
begin
  TBase.Announce;
end.
"#,
    );
    assert_eq!(out, vec!["HelperStaticMethod"]);
}

#[test]
fn test_record_helper_on_literal_values() {
    let out = run_pascal(
        r#"
program Test;
type TIntHelper = record helper for Integer
  function DoubleVal: Integer;
end;
function TIntHelper.DoubleVal: Integer; begin Result := Self * 2; end;
var res: Integer;
begin
  res := 50;
  WriteLn(res.DoubleVal);
end.
"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_record_helper_char_methods() {
    let out = run_pascal(
        r#"
program Test;
type TCharHelper = record helper for Char
  function IsDigitChar: Boolean;
end;
function TCharHelper.IsDigitChar: Boolean;
begin
  Result := (Self >= '0') and (Self <= '9');
end;
var c: Char;
begin
  c := '8';
  WriteLn(c.IsDigitChar);
  c := 'X';
  WriteLn(c.IsDigitChar);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_class_helper_inheritance_hierarchy() {
    let out = run_pascal(
        r#"
program Test;
type TParent = class public PName: String; end;
type TChild = class(TParent) public CCode: Integer; end;
type TParentHelper = class helper for TParent
  function GetPInfo: String;
end;
function TParentHelper.GetPInfo: String; begin Result := 'P:' + Self.PName; end;
var c: TChild;
begin
  c := TChild.Create;
  c.PName := 'ParentData';
  WriteLn(c.GetPInfo);
  c.Free;
end.
"#,
    );
    assert_eq!(out, vec!["P:ParentData"]);
}

#[test]
fn test_record_helper_returning_string() {
    let out = run_pascal(
        r#"
program Test;
type TIntFormatHelper = record helper for Integer
  function ToPaddedString(len: Integer): String;
end;
function TIntFormatHelper.ToPaddedString(len: Integer): String;
begin
  Result := Self.ToString;
  while Length(Result) < len do
    Result := '0' + Result;
end;
var id: Integer;
begin
  id := 42;
  WriteLn(id.ToPaddedString(5));
end.
"#,
    );
    assert_eq!(out, vec!["00042"]);
}
