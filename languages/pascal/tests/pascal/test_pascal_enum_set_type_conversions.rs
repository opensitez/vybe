use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 95: Enum & Set Type Reflection & Bitmask Conversions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_enum_getenumname_rtti() {
    let out = run_pascal(
        r#"
program Test;
uses TypInfo;
type TColor = (cRed, cGreen, cBlue);
begin
  WriteLn(GetEnumName(TypeInfo(TColor), Ord(cRed)));
  WriteLn(GetEnumName(TypeInfo(TColor), Ord(cGreen)));
  WriteLn(GetEnumName(TypeInfo(TColor), Ord(cBlue)));
end.
"#,
    );
    assert_eq!(out, vec!["cRed", "cGreen", "cBlue"]);
}

#[test]
fn test_enum_getenumvalue_rtti() {
    let out = run_pascal(
        r#"
program Test;
uses TypInfo;
type TColor = (cRed, cGreen, cBlue);
var val: Integer; c: TColor;
begin
  val := GetEnumValue(TypeInfo(TColor), 'cGreen');
  c := TColor(val);
  WriteLn(Ord(c));
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_enum_custom_ordinal_values() {
    let out = run_pascal(
        r#"
program Test;
type TCode = (cOne = 1, cTen = 10, cHundred = 100);
begin
  WriteLn(Ord(cOne));
  WriteLn(Ord(cTen));
  WriteLn(Ord(cHundred));
end.
"#,
    );
    assert_eq!(out, vec!["1", "10", "100"]);
}

#[test]
fn test_enum_low_high_bounds() {
    let out = run_pascal(
        r#"
program Test;
type TFruit = (fApple = 5, fOrange = 20);
begin
  WriteLn(Ord(Low(TFruit)));
  WriteLn(Ord(High(TFruit)));
end.
"#,
    );
    assert_eq!(out, vec!["5", "20"]);
}

#[test]
fn test_set_bitmask_byte_conversion() {
    let out = run_pascal(
        r#"
program Test;
type TBit = (b0, b1, b2, b3, b4, b5, b6, b7);
type TBits = set of TBit;
var s: TBits; b: Byte;
begin
  s := [b0, b3, b7]; // 1 + 8 + 128 = 137 ($89)
  b := PByte(@s)^;
  WriteLn(b);
end.
"#,
    );
    assert_eq!(out, vec!["137"]);
}

#[test]
fn test_byte_to_set_bitmask_cast() {
    let out = run_pascal(
        r#"
program Test;
type TBit = (b0, b1, b2, b3, b4, b5, b6, b7);
type TBits = set of TBit;
var s: TBits; b: Byte;
begin
  b := 5; // 1 + 4 = b0, b2
  PByte(@s)^ := b;
  WriteLn(b0 in s);
  WriteLn(b1 in s);
  WriteLn(b2 in s);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False", "True"]);
}

#[test]
fn test_set_settostring_rtti() {
    let out = run_pascal(
        r#"
program Test;
uses TypInfo;
type TFeature = (ftA, ftB, ftC);
type TFeatures = set of TFeature;
var f: TFeatures;
begin
  f := [ftA, ftC];
  WriteLn(SetToString(TypeInfo(TFeatures), @f, True));
end.
"#,
    );
    assert_eq!(out, vec!["[ftA, ftC]"]);
}

#[test]
fn test_set_stringtoset_rtti() {
    let out = run_pascal(
        r#"
program Test;
uses TypInfo;
type TFeature = (ftA, ftB, ftC);
type TFeatures = set of TFeature;
var f: TFeatures;
begin
  StringToSet(TypeInfo(TFeatures), '[ftB]', @f);
  WriteLn(ftB in f);
  WriteLn(ftA in f);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_enum_succ_pred_custom_ordinals() {
    let out = run_pascal(
        r#"
program Test;
type TMode = (mAlpha, mBeta, mGamma);
begin
  WriteLn(GetEnumName(TypeInfo(TMode), Ord(Succ(mAlpha))));
  WriteLn(GetEnumName(TypeInfo(TMode), Ord(Pred(mGamma))));
end.
"#,
    );
    assert_eq!(out, vec!["mBeta", "mBeta"]);
}

#[test]
fn test_enum_to_integer_explicit_cast() {
    let out = run_pascal(
        r#"
program Test;
type TLevel = (lvlLow = 10, lvlHigh = 50);
var l: TLevel; i: Integer;
begin
  l := lvlHigh;
  i := Integer(l);
  WriteLn(i);
end.
"#,
    );
    assert_eq!(out, vec!["50"]);
}

#[test]
fn test_integer_to_enum_explicit_cast() {
    let out = run_pascal(
        r#"
program Test;
uses TypInfo;
type TLevel = (lvlLow = 10, lvlHigh = 50);
var l: TLevel;
begin
  l := TLevel(10);
  WriteLn(GetEnumName(TypeInfo(TLevel), Ord(l)));
end.
"#,
    );
    assert_eq!(out, vec!["lvlLow"]);
}

#[test]
fn test_set_cardinal_bitmask_conversion() {
    let out = run_pascal(
        r#"
program Test;
type TBit32 = 0..31;
type TSet32 = set of TBit32;
var s: TSet32; card: Cardinal;
begin
  s := [0, 31]; // 1 + 2^31
  card := PCardinal(@s)^;
  WriteLn(card <> 0);
  WriteLn(0 in s);
  WriteLn(31 in s);
end.
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_set_population_count_elements() {
    let out = run_pascal(
        r#"
program Test;
type TOption = (opt1, opt2, opt3, opt4);
type TOptions = set of TOption;

function SetElemCount(const s: TOptions): Integer;
var opt: TOption;
begin
  Result := 0;
  for opt := Low(TOption) to High(TOption) do
    if opt in s then Inc(Result);
end;

var opts: TOptions;
begin
  opts := [opt1, opt3, opt4];
  WriteLn(SetElemCount(opts));
end.
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_enum_invalid_string_getenumvalue_returns_minus_one() {
    let out = run_pascal(
        r#"
program Test;
uses TypInfo;
type TColor = (cRed, cGreen);
begin
  WriteLn(GetEnumValue(TypeInfo(TColor), 'cNonExistent'));
end.
"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn test_enum_typeinfo_kind_check() {
    let out = run_pascal(
        r#"
program Test;
uses TypInfo;
type TChoice = (chYes, chNo);
begin
  WriteLn(TypeInfo(TChoice)^.Kind = tkEnumeration);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_set_typeinfo_kind_check() {
    let out = run_pascal(
        r#"
program Test;
uses TypInfo;
type TChoice = (chYes, chNo);
type TChoices = set of TChoice;
begin
  WriteLn(TypeInfo(TChoices)^.Kind = tkSet);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_enum_for_in_loop_reflection() {
    let out = run_pascal(
        r#"
program Test;
uses TypInfo;
type TLetter = (lA, lB, lC);
var i: Integer;
begin
  for i := Ord(Low(TLetter)) to Ord(High(TLetter)) do
    WriteLn(GetEnumName(TypeInfo(TLetter), i));
end.
"#,
    );
    assert_eq!(out, vec!["lA", "lB", "lC"]);
}

#[test]
fn test_set_empty_set_bitmask_is_zero() {
    let out = run_pascal(
        r#"
program Test;
type TBit = (b0, b1);
type TBits = set of TBit;
var s: TBits; b: Byte;
begin
  s := [];
  b := PByte(@s)^;
  WriteLn(b = 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_set_full_set_bitmask() {
    let out = run_pascal(
        r#"
program Test;
type TBit = (b0, b1, b2, b3);
type TBits = set of TBit;
var s: TBits; b: Byte;
begin
  s := [b0, b1, b2, b3]; // 1 + 2 + 4 + 8 = 15
  b := PByte(@s)^;
  WriteLn(b);
end.
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_enum_subrange_typeinfo() {
    let out = run_pascal(
        r#"
program Test;
uses TypInfo;
type TColor = (cRed, cGreen, cBlue, cYellow);
type TSubColor = cGreen..cBlue;
begin
  WriteLn(GetEnumName(TypeInfo(TColor), Ord(Low(TSubColor))));
  WriteLn(GetEnumName(TypeInfo(TColor), Ord(High(TSubColor))));
end.
"#,
    );
    assert_eq!(out, vec!["cGreen", "cBlue"]);
}
