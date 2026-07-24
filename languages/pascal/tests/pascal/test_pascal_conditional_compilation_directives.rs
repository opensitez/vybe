use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 88: Conditional Compilation Directives & Macro Expansion
// ═══════════════════════════════════════════════════════════

#[test]
fn test_cond_ifdef_endif_defined() {
    let out = run_pascal(r#"
program Test;
{$DEFINE FEATURE_ENABLED}
begin
  {$IFDEF FEATURE_ENABLED}
  WriteLn('FeatureOn');
  {$ELSE}
  WriteLn('FeatureOff');
  {$ENDIF}
end.
"#);
    assert_eq!(out, vec!["FeatureOn"]);
}

#[test]
fn test_cond_ifndef_endif_not_defined() {
    let out = run_pascal(r#"
program Test;
{$UNDEF FEATURE_DEBUG}
begin
  {$IFNDEF FEATURE_DEBUG}
  WriteLn('DebugDisabled');
  {$ENDIF}
end.
"#);
    assert_eq!(out, vec!["DebugDisabled"]);
}

#[test]
fn test_cond_if_defined_expression() {
    let out = run_pascal(r#"
program Test;
{$DEFINE MODE_RELEASE}
begin
  {$IF DEFINED(MODE_RELEASE)}
  WriteLn('ReleaseMode');
  {$ELSE}
  WriteLn('OtherMode');
  {$ENDIF}
end.
"#);
    assert_eq!(out, vec!["ReleaseMode"]);
}

#[test]
fn test_cond_elseif_branching() {
    let out = run_pascal(r#"
program Test;
{$DEFINE TARGET_MACOS}
begin
  {$IFDEF TARGET_WINDOWS}
  WriteLn('WinOS');
  {$ELSEIF DEFINED(TARGET_MACOS)}
  WriteLn('MacOS');
  {$ELSE}
  WriteLn('LinuxOS');
  {$ENDIF}
end.
"#);
    assert_eq!(out, vec!["MacOS"]);
}

#[test]
fn test_cond_macro_expansion() {
    let out = run_pascal(r#"
program Test;
{$MACRO ON}
{$DEFINE APP_NAME := 'PascalApp'}
begin
  WriteLn(APP_NAME);
end.
"#);
    assert_eq!(out, vec!["PascalApp"]);
}

#[test]
fn test_cond_rangechecks_directive() {
    let out = run_pascal(r#"
program Test;
{$R+} // Range checks ON
var val: 1..10;
begin
  val := 5;
  WriteLn(val);
end.
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_cond_overflowchecks_directive() {
    let out = run_pascal(r#"
program Test;
{$Q+} // Overflow checks ON
var a, b: Integer;
begin
  a := 10; b := 20;
  WriteLn(a + b);
end.
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_cond_assertions_directive() {
    let out = run_pascal(r#"
program Test;
{$C+} // Assertions ON
begin
  Assert(10 > 2);
  WriteLn('AssertOK');
end.
"#);
    assert_eq!(out, vec!["AssertOK"]);
}

#[test]
fn test_cond_mode_delphi_directive() {
    let out = run_pascal(r#"
{$MODE DELPHI}
program Test;
type TRec = record Val: Integer; end;
var r: TRec;
begin
  r.Val := 99;
  WriteLn(r.Val);
end.
"#);
    assert_eq!(out, vec!["99"]);
}

#[test]
fn test_cond_mode_objfpc_directive() {
    let out = run_pascal(r#"
{$MODE OBJFPC}
program Test;
var val: Integer;
begin
  val := 42;
  WriteLn(val);
end.
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_cond_align_directive() {
    let out = run_pascal(r#"
program Test;
{$ALIGN 4}
type TAlignedRec = record
  b: Byte;
  i: Integer;
end;
begin
  WriteLn(SizeOf(TAlignedRec) >= 8);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_cond_packrecords_1_directive() {
    let out = run_pascal(r#"
program Test;
{$PACKRECORDS 1}
type TPackedRec = record
  b: Byte;
  i: Integer;
end;
begin
  WriteLn(SizeOf(TPackedRec) = 5);
end.
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_cond_nested_ifdef_blocks() {
    let out = run_pascal(r#"
program Test;
{$DEFINE OPT_A}
{$DEFINE OPT_B}
begin
  {$IFDEF OPT_A}
    {$IFDEF OPT_B}
    WriteLn('OptA_And_OptB');
    {$ENDIF}
  {$ENDIF}
end.
"#);
    assert_eq!(out, vec!["OptA_And_OptB"]);
}

#[test]
fn test_cond_boolean_expression_defined() {
    let out = run_pascal(r#"
program Test;
{$DEFINE FLAG1}
{$DEFINE FLAG2}
begin
  {$IF DEFINED(FLAG1) AND DEFINED(FLAG2)}
  WriteLn('BothFlagsActive');
  {$ENDIF}
end.
"#);
    assert_eq!(out, vec!["BothFlagsActive"]);
}

#[test]
fn test_cond_boolean_or_expression() {
    let out = run_pascal(r#"
program Test;
{$DEFINE FLAG_A}
begin
  {$IF DEFINED(FLAG_A) OR DEFINED(FLAG_B)}
  WriteLn('AtLeastOneFlagActive');
  {$ENDIF}
end.
"#);
    assert_eq!(out, vec!["AtLeastOneFlagActive"]);
}

#[test]
fn test_cond_undef_deactivates_symbol() {
    let out = run_pascal(r#"
program Test;
{$DEFINE TEMP_FEATURE}
{$UNDEF TEMP_FEATURE}
begin
  {$IFDEF TEMP_FEATURE}
  WriteLn('Active');
  {$ELSE}
  WriteLn('Deactivated');
  {$ENDIF}
end.
"#);
    assert_eq!(out, vec!["Deactivated"]);
}

#[test]
fn test_cond_hints_warnings_directives() {
    let out = run_pascal(r#"
program Test;
{$HINTS OFF}
{$WARNINGS OFF}
var unusedVar: Integer;
begin
  WriteLn('DirectivesProcessed');
end.
"#);
    assert_eq!(out, vec!["DirectivesProcessed"]);
}

#[test]
fn test_cond_typedaddress_directive() {
    let out = run_pascal(r#"
program Test;
{$T+} // Typed @ operator ON
var x: Integer; p: PInteger;
begin
  p := @x;
  p^ := 100;
  WriteLn(x);
end.
"#);
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_cond_pointermath_directive() {
    let out = run_pascal(r#"
program Test;
{$POINTERMATH ON}
var arr: array[0..2] of Integer; p: PInteger;
begin
  arr[0] := 10; arr[1] := 20; arr[2] := 30;
  p := @arr[0];
  WriteLn(p[1]);
end.
"#);
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_cond_iochecks_directive() {
    let out = run_pascal(r#"
program Test;
{$I-} // I/O checks OFF
var f: TextFile;
begin
  AssignFile(f, 'non_existent_file_xyz_123.txt');
  Reset(f);
  WriteLn(IOResult <> 0);
  {$I+} // I/O checks back ON
end.
"#);
    assert_eq!(out, vec!["True"]);
}
