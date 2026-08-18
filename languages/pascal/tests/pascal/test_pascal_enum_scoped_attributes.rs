use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 3: Scoped Enumerations & Enum Attributes
// ═══════════════════════════════════════════════════════════

#[test]
fn test_enum_basic_declaration_and_ord() {
    let out = run_pascal(
        r#"
program Test;
type TStatus = (Pending, Active, Completed, Cancelled);
var s: TStatus;
begin
  s := Active;
  WriteLn(Ord(s));
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_enum_explicit_ordinal_assignments() {
    let out = run_pascal(
        r#"
program Test;
type TFlags = (FlagA = 1, FlagB = 4, FlagC = 16);
begin
  WriteLn(Ord(FlagA));
  WriteLn(Ord(FlagB));
  WriteLn(Ord(FlagC));
end.
"#,
    );
    assert_eq!(out, vec!["1", "4", "16"]);
}

#[test]
fn test_enum_high_and_low_bounds() {
    let out = run_pascal(
        r#"
program Test;
type TPriority = (LowPri, MedPri, HighPri, CriticalPri);
begin
  WriteLn(Ord(Low(TPriority)));
  WriteLn(Ord(High(TPriority)));
end.
"#,
    );
    assert_eq!(out, vec!["0", "3"]);
}

#[test]
fn test_enum_pred_and_succ_operations() {
    let out = run_pascal(
        r#"
program Test;
type TDirection = (North, East, South, West);
var d: TDirection;
begin
  d := East;
  WriteLn(Ord(Pred(d)));
  WriteLn(Ord(Succ(d)));
end.
"#,
    );
    assert_eq!(out, vec!["0", "2"]);
}

#[test]
fn test_enum_typecast_from_integer() {
    let out = run_pascal(
        r#"
program Test;
type TSeason = (Spring, Summer, Autumn, Winter);
var s: TSeason;
begin
  s := TSeason(2);
  WriteLn(Ord(s));
end.
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_enum_comparison_operators() {
    let out = run_pascal(
        r#"
program Test;
type TLevel = (Novice, Intermediate, Expert, Master);
begin
  WriteLn(Novice < Expert);
  WriteLn(Master > Intermediate);
  WriteLn(Novice = Novice);
  WriteLn(Novice <> Expert);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "TRUE", "TRUE", "TRUE"]);
}

#[test]
fn test_enum_case_statement_branching() {
    let out = run_pascal(
        r#"
program Test;
type TState = (StateOff, StateOn, StateStandby);
procedure CheckState(s: TState);
begin
  case s of
    StateOff: WriteLn('Off');
    StateOn: WriteLn('On');
    StateStandby: WriteLn('Standby');
  end;
end;
begin
  CheckState(StateOn);
end.
"#,
    );
    assert_eq!(out, vec!["On"]);
}

#[test]
fn test_enum_for_loop_iteration() {
    let out = run_pascal(
        r#"
program Test;
type TWeekday = (Mon, Tue, Wed, Thu, Fri);
var d: TWeekday;
    sum: Integer;
begin
  sum := 0;
  for d := Low(TWeekday) to High(TWeekday) do
    sum := sum + Ord(d);
  WriteLn(sum);
end.
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_enum_as_array_indexer() {
    let out = run_pascal(
        r#"
program Test;
type TColor = (Red, Green, Blue);
type TColorMap = array[TColor] of String;
var map: TColorMap;
begin
  map[Red] := '#FF0000';
  map[Green] := '#00FF00';
  map[Blue] := '#0000FF';
  WriteLn(map[Green]);
end.
"#,
    );
    assert_eq!(out, vec!["#00FF00"]);
}

#[test]
fn test_enum_in_record_structure() {
    let out = run_pascal(
        r#"
program Test;
type TRole = (Guest, User, Admin);
type TUserRec = record
  Username: String;
  Role: TRole;
end;
var u: TUserRec;
begin
  u.Username := 'Bob';
  u.Role := Admin;
  WriteLn(u.Username);
  WriteLn(Ord(u.Role));
end.
"#,
    );
    assert_eq!(out, vec!["Bob", "2"]);
}

#[test]
fn test_enum_set_creation() {
    let out = run_pascal(
        r#"
program Test;
type TPermission = (ReadPerm, WritePerm, ExecPerm);
type TPermissions = set of TPermission;
var p: TPermissions;
begin
  p := [ReadPerm, ExecPerm];
  WriteLn(ReadPerm in p);
  WriteLn(WritePerm in p);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_enum_inc_dec_mutations() {
    let out = run_pascal(
        r#"
program Test;
type TStep = (Step1, Step2, Step3, Step4);
var s: TStep;
begin
  s := Step1;
  Inc(s);
  WriteLn(Ord(s));
  Inc(s, 2);
  WriteLn(Ord(s));
  Dec(s);
  WriteLn(Ord(s));
end.
"#,
    );
    assert_eq!(out, vec!["1", "3", "2"]);
}

#[test]
fn test_enum_constant_array_lookup() {
    let out = run_pascal(
        r#"
program Test;
type TMonth = (Jan, Feb, Mar);
const MonthNames: array[TMonth] of String = ('January', 'February', 'March');
begin
  WriteLn(MonthNames[Feb]);
end.
"#,
    );
    assert_eq!(out, vec!["February"]);
}

#[test]
fn test_enum_function_return_value() {
    let out = run_pascal(
        r#"
program Test;
type TMode = (ModeA, ModeB);
function GetNextMode(m: TMode): TMode;
begin
  if m = ModeA then Result := ModeB else Result := ModeA;
end;
begin
  WriteLn(Ord(GetNextMode(ModeA)));
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_enum_subrange_derivation() {
    let out = run_pascal(
        r#"
program Test;
type TFull = (E1, E2, E3, E4, E5);
type TPart = E2..E4;
var p: TPart;
begin
  p := E3;
  WriteLn(Ord(p));
end.
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_enum_assigned_gaps() {
    let out = run_pascal(
        r#"
program Test;
type TCode = (CodeA = 10, CodeB = 20, CodeC = 30);
var c: TCode;
begin
  c := CodeB;
  WriteLn(Ord(c));
end.
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_enum_scoped_directive_access() {
    let out = run_pascal(
        r#"
program Test;
{$SCOPEDENUMS ON}
type TKind = (Alpha, Beta, Gamma);
{$SCOPEDENUMS OFF}
var k: TKind;
begin
  k := TKind.Beta;
  WriteLn(Ord(k));
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_enum_size_of() {
    let out = run_pascal(
        r#"
program Test;
type TSmallEnum = (A, B, C);
begin
  WriteLn(SizeOf(TSmallEnum) > 0);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_enum_default_value_in_var() {
    let out = run_pascal(
        r#"
program Test;
type TState = (Init, Running, Stopped);
var s: TState;
begin
  s := Low(TState);
  WriteLn(Ord(s));
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_enum_bitwise_flags_combination() {
    let out = run_pascal(
        r#"
program Test;
type TFlagBit = (Bit0 = 1, Bit1 = 2, Bit2 = 4, Bit3 = 8);
var combined: Integer;
begin
  combined := Ord(Bit0) or Ord(Bit2);
  WriteLn(combined);
end.
"#,
    );
    assert_eq!(out, vec!["5"]);
}
