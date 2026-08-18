use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 90: Unit Initialization & Finalization Sections
// ═══════════════════════════════════════════════════════════

#[test]
fn test_unit_initialization_execution_order() {
    let out = run_pascal(
        r#"
unit TestUnitInit;
interface
  procedure Dummy;
implementation
procedure Dummy; begin end;
initialization
  WriteLn('UnitInitialized');
finalization
  WriteLn('UnitFinalized');
end.

program Test;
uses TestUnitInit;
begin
  WriteLn('MainProgramBody');
end.
"#,
    );
    assert_eq!(
        out,
        vec!["UnitInitialized", "MainProgramBody", "UnitFinalized"]
    );
}

#[test]
fn test_unit_initialization_global_variable() {
    let out = run_pascal(
        r#"
unit GlobalInitUnit;
interface
  var GlobalCounter: Integer;
implementation
initialization
  GlobalCounter := 100;
end.

program Test;
uses GlobalInitUnit;
begin
  WriteLn(GlobalCounter);
end.
"#,
    );
    assert_eq!(out, vec!["100"]);
}

#[test]
fn test_unit_finalization_cleanup_variable() {
    let out = run_pascal(
        r#"
unit CleanupUnit;
interface
  procedure Touch;
implementation
procedure Touch; begin WriteLn('TouchWork'); end;
initialization
  WriteLn('InitCleanupUnit');
finalization
  WriteLn('FinalizeCleanupUnit');
end.

program Test;
uses CleanupUnit;
begin
  Touch;
end.
"#,
    );
    assert_eq!(
        out,
        vec!["InitCleanupUnit", "TouchWork", "FinalizeCleanupUnit"]
    );
}

#[test]
fn test_unit_initialization_singleton_instantiation() {
    let out = run_pascal(
        r#"
unit SingletonUnit;
interface
  type TSingleton = class
    public procedure Speak;
  end;
  var SingletonInstance: TSingleton;
implementation
procedure TSingleton.Speak; begin WriteLn('SingletonActive'); end;
initialization
  SingletonInstance := TSingleton.Create;
finalization
  SingletonInstance.Free;
  WriteLn('SingletonFreed');
end.

program Test;
uses SingletonUnit;
begin
  SingletonInstance.Speak;
end.
"#,
    );
    assert_eq!(out, vec!["SingletonActive", "SingletonFreed"]);
}

#[test]
fn test_unit_dependencies_initialization_order() {
    let out = run_pascal(
        r#"
unit UnitA;
interface procedure ProcA;
implementation
procedure ProcA; begin end;
initialization
  WriteLn('InitA');
finalization
  WriteLn('FinalA');
end.

unit UnitB;
interface uses UnitA; procedure ProcB;
implementation
procedure ProcB; begin ProcA; end;
initialization
  WriteLn('InitB');
finalization
  WriteLn('FinalB');
end.

program Test;
uses UnitB;
begin
  WriteLn('MainProgram');
end.
"#,
    );
    assert_eq!(
        out,
        vec!["InitA", "InitB", "MainProgram", "FinalB", "FinalA"]
    );
}

#[test]
fn test_unit_initialization_stringlist_population() {
    let out = run_pascal(
        r#"
unit ConfigUnit;
interface
  uses Classes;
  var AppConfig: TStringList;
implementation
initialization
  AppConfig := TStringList.Create;
  AppConfig.Add('Key=Value');
finalization
  AppConfig.Free;
end.

program Test;
uses ConfigUnit;
begin
  WriteLn(AppConfig.Values['Key']);
end.
"#,
    );
    assert_eq!(out, vec!["Value"]);
}

#[test]
fn test_unit_initialization_with_try_except() {
    let out = run_pascal(
        r#"
unit SafeInitUnit;
interface
  var SafeInitSuccess: Boolean;
implementation
uses SysUtils;
initialization
  try
    SafeInitSuccess := True;
  except
    SafeInitSuccess := False;
  end;
end.

program Test;
uses SafeInitUnit;
begin
  WriteLn(SafeInitSuccess);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_unit_initialization_nested_routine_call() {
    let out = run_pascal(
        r#"
unit NestedInitUnit;
interface
  var InitValue: Integer;
implementation
function ComputeInitVal: Integer; begin Result := 42; end;
initialization
  InitValue := ComputeInitVal;
end.

program Test;
uses NestedInitUnit;
begin
  WriteLn(InitValue);
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_unit_finalization_without_initialization() {
    let out = run_pascal(
        r#"
unit FinalOnlyUnit;
interface
  procedure Run;
implementation
procedure Run; begin WriteLn('InsideRun'); end;
finalization
  WriteLn('FinalOnlyExecuted');
end.

program Test;
uses FinalOnlyUnit;
begin
  Run;
end.
"#,
    );
    assert_eq!(out, vec!["InsideRun", "FinalOnlyExecuted"]);
}

#[test]
fn test_unit_initialization_without_finalization() {
    let out = run_pascal(
        r#"
unit InitOnlyUnit;
interface
  var Flag: Boolean;
implementation
initialization
  Flag := True;
end.

program Test;
uses InitOnlyUnit;
begin
  WriteLn(Flag);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_unit_initialization_array_setup() {
    let out = run_pascal(
        r#"
unit LookupUnit;
interface
  var Squares: array[0..3] of Integer;
implementation
var i: Integer;
initialization
  for i := 0 to 3 do
    Squares[i] := i * i;
end.

program Test;
uses LookupUnit;
begin
  WriteLn(Squares[3]);
end.
"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn test_unit_initialization_proc_pointer_assignment() {
    let out = run_pascal(
        r#"
unit CallbackUnit;
interface
  type TCallback = procedure(const s: String);
  var GlobalCallback: TCallback;
implementation
procedure DefaultCallback(const s: String);
begin
  WriteLn('DefaultCB:' + s);
end;
initialization
  GlobalCallback := DefaultCallback;
end.

program Test;
uses CallbackUnit;
begin
  GlobalCallback('TestMsg');
end.
"#,
    );
    assert_eq!(out, vec!["DefaultCB:TestMsg"]);
}

#[test]
fn test_unit_finalization_with_try_finally() {
    let out = run_pascal(
        r#"
unit SafeFinalUnit;
interface
  procedure DoWork;
implementation
procedure DoWork; begin end;
finalization
  try
    WriteLn('FinalizeWork');
  finally
    WriteLn('FinalizeFinally');
  end;
end.

program Test;
uses SafeFinalUnit;
begin
  DoWork;
end.
"#,
    );
    assert_eq!(out, vec!["FinalizeWork", "FinalizeFinally"]);
}

#[test]
fn test_unit_initialization_rtti_registration() {
    let out = run_pascal(
        r#"
unit RegUnit;
interface
  var RegisteredName: String;
implementation
initialization
  RegisteredName := 'RegUnitClass';
end.

program Test;
uses RegUnit;
begin
  WriteLn(RegisteredName);
end.
"#,
    );
    assert_eq!(out, vec!["RegUnitClass"]);
}

#[test]
fn test_unit_initialization_flag_toggle() {
    let out = run_pascal(
        r#"
unit FlagUnit;
interface
  var IsReady: Boolean;
implementation
initialization
  IsReady := True;
finalization
  IsReady := False;
end.

program Test;
uses FlagUnit;
begin
  WriteLn(IsReady);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_unit_initialization_math_constant_setup() {
    let out = run_pascal(
        r#"
unit MathConstUnit;
interface
  var TwoPi: Double;
implementation
initialization
  TwoPi := 2.0 * 3.1415926535;
end.

program Test;
uses MathConstUnit;
begin
  WriteLn(TwoPi > 6.28);
end.
"#,
    );
    assert_eq!(out, vec!["TRUE"]);
}

#[test]
fn test_unit_initialization_record_setup() {
    let out = run_pascal(
        r#"
unit RecInitUnit;
interface
  type TPoint = record X, Y: Integer; end;
  var DefaultPoint: TPoint;
implementation
initialization
  DefaultPoint.X := 10;
  DefaultPoint.Y := 20;
end.

program Test;
uses RecInitUnit;
begin
  WriteLn(DefaultPoint.X.ToString + ',' + DefaultPoint.Y.ToString);
end.
"#,
    );
    assert_eq!(out, vec!["10,20"]);
}

#[test]
fn test_unit_finalization_decrements_counter() {
    let out = run_pascal(
        r#"
unit CounterUnit;
interface
  var Counter: Integer;
implementation
initialization
  Counter := 1;
finalization
  Dec(Counter);
  WriteLn('CounterFinalized:' + Counter.ToString);
end.

program Test;
uses CounterUnit;
begin
  WriteLn('CounterBody:' + Counter.ToString);
end.
"#,
    );
    assert_eq!(out, vec!["CounterBody:1", "CounterFinalized:0"]);
}

#[test]
fn test_unit_initialization_char_table() {
    let out = run_pascal(
        r#"
unit CharTableUnit;
interface
  var HexChars: array[0..15] of Char;
implementation
var i: Integer;
initialization
  for i := 0 to 9 do HexChars[i] := Chr(Ord('0') + i);
  for i := 10 to 15 do HexChars[i] := Chr(Ord('A') + i - 10);
end.

program Test;
uses CharTableUnit;
begin
  WriteLn(HexChars[10] + HexChars[15]);
end.
"#,
    );
    assert_eq!(out, vec!["AF"]);
}

#[test]
fn test_unit_initialization_interface_instance() {
    let out = run_pascal(
        r#"
unit IntfInitUnit;
interface
  type IService = interface
    ['{12345678-1234-1234-1234-123456789012}']
    procedure Execute;
  end;
  var ServiceRef: IService;
implementation
type TServiceImpl = class(TInterfacedObject, IService)
  public procedure Execute;
end;
procedure TServiceImpl.Execute; begin WriteLn('InitServiceExecuted'); end;

initialization
  ServiceRef := TServiceImpl.Create;
finalization
  ServiceRef := nil;
end.

program Test;
uses IntfInitUnit;
begin
  ServiceRef.Execute;
end.
"#,
    );
    assert_eq!(out, vec!["InitServiceExecuted"]);
}
