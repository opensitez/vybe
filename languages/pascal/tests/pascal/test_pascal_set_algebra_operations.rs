use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 4: Set Algebra & Set Operations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_set_empty_and_element_membership() {
    let out = run_pascal(r#"
program Test;
type TNumSet = set of 1..10;
var s: TNumSet;
begin
  s := [];
  WriteLn(5 in s);
  s := [5, 8];
  WriteLn(5 in s);
  WriteLn(8 in s);
  WriteLn(1 in s);
end.
"#);
    assert_eq!(out, vec!["False", "True", "True", "False"]);
}

#[test]
fn test_set_union_operator() {
    let out = run_pascal(r#"
program Test;
type TCharSet = set of Char;
var s1, s2, s3: TCharSet;
begin
  s1 := ['A', 'B'];
  s2 := ['B', 'C'];
  s3 := s1 + s2;
  WriteLn('A' in s3);
  WriteLn('B' in s3);
  WriteLn('C' in s3);
end.
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_set_intersection_operator() {
    let out = run_pascal(r#"
program Test;
type TNumSet = set of 1..20;
var s1, s2, s3: TNumSet;
begin
  s1 := [1, 2, 3, 4, 5];
  s2 := [3, 4, 5, 6, 7];
  s3 := s1 * s2;
  WriteLn(1 in s3);
  WriteLn(3 in s3);
  WriteLn(5 in s3);
  WriteLn(7 in s3);
end.
"#);
    assert_eq!(out, vec!["False", "True", "True", "False"]);
}

#[test]
fn test_set_difference_operator() {
    let out = run_pascal(r#"
program Test;
type TNumSet = set of 1..10;
var s1, s2, s3: TNumSet;
begin
  s1 := [1, 2, 3, 4, 5];
  s2 := [2, 4];
  s3 := s1 - s2;
  WriteLn(1 in s3);
  WriteLn(2 in s3);
  WriteLn(3 in s3);
  WriteLn(4 in s3);
end.
"#);
    assert_eq!(out, vec!["True", "False", "True", "False"]);
}

#[test]
fn test_set_symmetric_difference_operator() {
    let out = run_pascal(r#"
program Test;
type TNumSet = set of 1..10;
var s1, s2, s3: TNumSet;
begin
  s1 := [1, 2, 3];
  s2 := [2, 3, 4];
  s3 := s1 >< s2;
  WriteLn(1 in s3);
  WriteLn(2 in s3);
  WriteLn(3 in s3);
  WriteLn(4 in s3);
end.
"#);
    assert_eq!(out, vec!["True", "False", "False", "True"]);
}

#[test]
fn test_set_include_procedure() {
    let out = run_pascal(r#"
program Test;
type TNumSet = set of 1..10;
var s: TNumSet;
begin
  s := [1, 2];
  Include(s, 3);
  WriteLn(3 in s);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_set_exclude_procedure() {
    let out = run_pascal(r#"
program Test;
type TNumSet = set of 1..10;
var s: TNumSet;
begin
  s := [1, 2, 3];
  Exclude(s, 2);
  WriteLn(2 in s);
  WriteLn(1 in s);
  WriteLn(3 in s);
end.
"#);
    assert_eq!(out, vec!["False", "True", "True"]);
}

#[test]
fn test_set_subset_comparison() {
    let out = run_pascal(r#"
program Test;
type TNumSet = set of 1..10;
var s1, s2: TNumSet;
begin
  s1 := [1, 2];
  s2 := [1, 2, 3, 4];
  WriteLn(s1 <= s2);
  WriteLn(s2 <= s1);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_set_superset_comparison() {
    let out = run_pascal(r#"
program Test;
type TNumSet = set of 1..10;
var s1, s2: TNumSet;
begin
  s1 := [1, 2, 3, 4];
  s2 := [1, 2];
  WriteLn(s1 >= s2);
  WriteLn(s2 >= s1);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_set_equality_and_inequality() {
    let out = run_pascal(r#"
program Test;
type TNumSet = set of 1..10;
var s1, s2, s3: TNumSet;
begin
  s1 := [1, 3, 5];
  s2 := [5, 1, 3];
  s3 := [1, 3];
  WriteLn(s1 = s2);
  WriteLn(s1 <> s3);
end.
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_set_range_literal_construction() {
    let out = run_pascal(r#"
program Test;
type TCharSet = set of Char;
var vowels: TCharSet;
begin
  vowels := ['a'..'e'];
  WriteLn('a' in vowels);
  WriteLn('c' in vowels);
  WriteLn('f' in vowels);
end.
"#);
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_set_enum_based_set() {
    let out = run_pascal(r#"
program Test;
type TStyle = (fsBold, fsItalic, fsUnderline, fsStrikeOut);
type TFontStyle = set of TStyle;
var font: TFontStyle;
begin
  font := [fsBold, fsItalic];
  WriteLn(fsBold in font);
  WriteLn(fsUnderline in font);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_set_parameter_passing() {
    let out = run_pascal(r#"
program Test;
type TDigitSet = set of 0..9;
function CountDigits(s: TDigitSet): Integer;
var i, count: Integer;
begin
  count := 0;
  for i := 0 to 9 do
    if i in s then Inc(count);
  Result := count;
end;
begin
  WriteLn(CountDigits([1, 3, 5, 7, 9]));
end.
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_set_function_return_type() {
    let out = run_pascal(r#"
program Test;
type TAlphaSet = set of 'A'..'Z';
function GetVowels: TAlphaSet;
begin
  Result := ['A', 'E', 'I', 'O', 'U'];
end;
begin
  WriteLn('E' in GetVowels);
  WriteLn('B' in GetVowels);
end.
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_set_conditional_membership_branch() {
    let out = run_pascal(r#"
program Test;
type TCharSet = set of Char;
function IsHexDigit(c: Char): Boolean;
begin
  Result := c in ['0'..'9', 'A'..'F', 'a'..'f'];
end;
begin
  WriteLn(IsHexDigit('5'));
  WriteLn(IsHexDigit('B'));
  WriteLn(IsHexDigit('z'));
end.
"#);
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_set_reassignment_with_expressions() {
    let out = run_pascal(r#"
program Test;
type TNumSet = set of 1..10;
var s: TNumSet;
begin
  s := [1, 2] + [3, 4] - [2];
  WriteLn(1 in s);
  WriteLn(2 in s);
  WriteLn(3 in s);
  WriteLn(4 in s);
end.
"#);
    assert_eq!(out, vec!["True", "False", "True", "True"]);
}

#[test]
fn test_set_record_field_storage() {
    let out = run_pascal(r#"
program Test;
type TAccess = (ReadAcc, WriteAcc, ExecAcc);
type TAccessSet = set of TAccess;
type TFilePermission = record
  FileName: String;
  Access: TAccessSet;
end;
var fp: TFilePermission;
begin
  fp.FileName := 'data.bin';
  fp.Access := [ReadAcc, WriteAcc];
  WriteLn(fp.FileName);
  WriteLn(ReadAcc in fp.Access);
  WriteLn(ExecAcc in fp.Access);
end.
"#);
    assert_eq!(out, vec!["data.bin", "True", "False"]);
}

#[test]
fn test_set_chained_include_exclude() {
    let out = run_pascal(r#"
program Test;
type TNumSet = set of 1..10;
var s: TNumSet;
begin
  s := [];
  Include(s, 1);
  Include(s, 5);
  Include(s, 9);
  Exclude(s, 5);
  WriteLn(1 in s);
  WriteLn(5 in s);
  WriteLn(9 in s);
end.
"#);
    assert_eq!(out, vec!["True", "False", "True"]);
}

#[test]
fn test_set_subrange_variable_elements() {
    let out = run_pascal(r#"
program Test;
type TNumSet = set of 1..20;
var s: TNumSet;
    x, y: Integer;
begin
  x := 7;
  y := 14;
  s := [x, y];
  WriteLn(7 in s);
  WriteLn(14 in s);
  WriteLn(10 in s);
end.
"#);
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_set_all_elements_present() {
    let out = run_pascal(r#"
program Test;
type TByteRange = 1..4;
type TFullSet = set of TByteRange;
var s: TFullSet;
begin
  s := [1..4];
  WriteLn(1 in s);
  WriteLn(2 in s);
  WriteLn(3 in s);
  WriteLn(4 in s);
end.
"#);
    assert_eq!(out, vec!["True", "True", "True", "True"]);
}
