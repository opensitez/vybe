use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 57: Assertion Handling & Custom Assertion Hooks (Assert)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_assert_condition_true_passes() {
    let out = run_pascal(r#"
program Test;
begin
  Assert(10 > 5);
  WriteLn('AssertionPassed');
end.
"#);
    assert_eq!(out, vec!["AssertionPassed"]);
}

#[test]
fn test_assert_condition_false_raises_eassertionfailed() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  try
    Assert(10 < 5);
  except
    on E: EAssertionFailed do WriteLn('AssertFailedCaught');
  end;
end.
"#);
    assert_eq!(out, vec!["AssertFailedCaught"]);
}

#[test]
fn test_assert_custom_message() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  try
    Assert(2 + 2 = 5, 'MathIsBroken');
  except
    on E: Exception do WriteLn(E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["MathIsBroken"]);
}

#[test]
fn test_assert_assertions_off_directive() {
    let out = run_pascal(r#"
program Test;
{$C-} // Assertions OFF
begin
  Assert(10 < 5);
  WriteLn('AssertionsOffIgnoredFailure');
end.
"#);
    assert_eq!(out, vec!["AssertionsOffIgnoredFailure"]);
}

#[test]
fn test_assert_assertions_on_directive() {
    let out = run_pascal(r#"
program Test;
{$C+} // Assertions ON
uses SysUtils;
begin
  try
    Assert(1 = 2);
  except
    on E: EAssertionFailed do WriteLn('AssertionsOnTriggered');
  end;
end.
"#);
    assert_eq!(out, vec!["AssertionsOnTriggered"]);
}

#[test]
fn test_assert_pointer_not_nil() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var val: Integer; p: PInteger;
begin
  val := 42; p := @val;
  try
    Assert(p <> nil, 'PointerIsNull');
    WriteLn('PointerValid');
  except
    on E: Exception do WriteLn(E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["PointerValid"]);
}

#[test]
fn test_assert_string_not_empty() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure ProcessName(const name: String);
begin
  Assert(name <> '', 'NameCannotBeEmpty');
  WriteLn('Name:' + name);
end;
begin
  try
    ProcessName('Alice');
    ProcessName('');
  except
    on E: Exception do WriteLn('Caught:' + E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["Name:Alice", "Caught:NameCannotBeEmpty"]);
}

#[test]
fn test_custom_asserterrorproc_hook() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var HookTriggered: Boolean;

procedure CustomAssertHandler(const Message, Filename: String; LineNumber: Integer; ErrorAddr: Pointer);
begin
  HookTriggered := True;
  WriteLn('CustomHook:' + Message);
end;

var oldHandler: TAssertErrorProc;
begin
  HookTriggered := False;
  oldHandler := AssertErrorProc;
  AssertErrorProc := CustomAssertHandler;

  Assert(1 = 2, 'HookMessage');

  AssertErrorProc := oldHandler;
end.
"#);
    assert_eq!(out, vec!["CustomHook:HookMessage"]);
}

#[test]
fn test_assert_in_class_method() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type TAccount = class
  private FBalance: Integer;
  public procedure Deposit(amount: Integer);
end;
procedure TAccount.Deposit(amount: Integer);
begin
  Assert(amount > 0, 'DepositMustBePositive');
  Inc(FBalance, amount);
end;
var acc: TAccount;
begin
  acc := TAccount.Create;
  try
    acc.Deposit(100);
    WriteLn('Deposit100OK');
    acc.Deposit(-50);
  except
    on E: Exception do WriteLn('Caught:' + E.Message);
  end;
  acc.Free;
end.
"#);
    assert_eq!(out, vec!["Deposit100OK", "Caught:DepositMustBePositive"]);
}

#[test]
fn test_assert_in_record_method() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type TPoint = record
  X, Y: Integer;
  procedure SetPos(AX, AY: Integer);
end;
procedure TPoint.SetPos(AX, AY: Integer);
begin
  Assert((AX >= 0) and (AY >= 0), 'CoordsMustBeNonNegative');
  X := AX; Y := AY;
end;
var pt: TPoint;
begin
  try
    pt.SetPos(10, 20);
    WriteLn('PosValid');
    pt.SetPos(-1, 5);
  except
    on E: Exception do WriteLn('Caught:' + E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["PosValid", "Caught:CoordsMustBeNonNegative"]);
}

#[test]
fn test_assert_inside_loop() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var arr: array[0..2] of Integer; i: Integer;
begin
  arr[0] := 2; arr[1] := 4; arr[2] := -1;
  for i := 0 to 2 do
  begin
    try
      Assert(arr[i] >= 0, 'ElementNegative');
      WriteLn('ElementOK:' + arr[i].ToString);
    except
      on E: Exception do WriteLn('CaughtAtIndex:' + i.ToString);
    end;
  end;
end.
"#);
    assert_eq!(out, vec!["ElementOK:2", "ElementOK:4", "CaughtAtIndex:2"]);
}

#[test]
fn test_assert_assigned_object() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type TWidget = class end;
var w: TWidget;
begin
  w := nil;
  try
    Assert(Assigned(w), 'WidgetNotAssigned');
  except
    on E: Exception do WriteLn(E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["WidgetNotAssigned"]);
}

#[test]
fn test_assert_array_index_range() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure AccessIndex(idx: Integer);
begin
  Assert((idx >= 0) and (idx <= 4), 'IndexOutOfBounds');
  WriteLn('IndexOK');
end;
begin
  try
    AccessIndex(2);
    AccessIndex(10);
  except
    on E: Exception do WriteLn(E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["IndexOK", "IndexOutOfBounds"]);
}

#[test]
fn test_assert_recursive_invariant() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
function Factorial(n: Integer): Integer;
begin
  Assert(n >= 0, 'FactorialNegativeInput');
  if n <= 1 then Result := 1
  else Result := n * Factorial(n - 1);
end;
begin
  try
    WriteLn(Factorial(4));
    Factorial(-2);
  except
    on E: Exception do WriteLn(E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["24", "FactorialNegativeInput"]);
}

#[test]
fn test_assert_in_constructor() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
type TStrictItem = class
  constructor Create(code: Integer);
end;
constructor TStrictItem.Create(code: Integer);
begin
  Assert(code > 1000, 'CodeTooSmall');
end;
var item: TStrictItem;
begin
  try
    item := TStrictItem.Create(50);
  except
    on E: Exception do WriteLn('CtorCaught:' + E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["CtorCaught:CodeTooSmall"]);
}

#[test]
fn test_assert_boolean_logic_combination() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var a, b: Boolean;
begin
  a := True; b := False;
  try
    Assert(a and b, 'BothMustBeTrue');
  except
    on E: Exception do WriteLn(E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["BothMustBeTrue"]);
}

#[test]
fn test_assert_restoring_asserterrorproc() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
procedure DummyAssertHandler(const Message, Filename: String; LineNumber: Integer; ErrorAddr: Pointer);
begin
  WriteLn('DummyHandler:' + Message);
end;
var savedHandler: TAssertErrorProc;
begin
  savedHandler := AssertErrorProc;
  AssertErrorProc := DummyAssertHandler;
  Assert(1 = 2, 'Msg1');
  AssertErrorProc := savedHandler;
  try
    Assert(1 = 2, 'Msg2');
  except
    on E: EAssertionFailed do WriteLn('StandardHandler:' + E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["DummyHandler:Msg1", "StandardHandler:Msg2"]);
}

#[test]
fn test_assert_in_function_return_val() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
function GetSquare(x: Integer): Integer;
begin
  Result := x * x;
  Assert(Result >= 0, 'SquareCannotBeNegative');
end;
begin
  WriteLn(GetSquare(5));
end.
"#);
    assert_eq!(out, vec!["25"]);
}

#[test]
fn test_assert_count_tracker_in_custom_hook() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
var assertCount: Integer;

procedure CountingAssertHandler(const Message, Filename: String; LineNumber: Integer; ErrorAddr: Pointer);
begin
  Inc(assertCount);
end;

var oldHandler: TAssertErrorProc;
begin
  assertCount := 0;
  oldHandler := AssertErrorProc;
  AssertErrorProc := CountingAssertHandler;

  Assert(1 = 2, 'Fail1');
  Assert(2 = 3, 'Fail2');

  AssertErrorProc := oldHandler;
  WriteLn(assertCount);
end.
"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_assert_multiline_custom_message() {
    let out = run_pascal(r#"
program Test;
uses SysUtils;
begin
  try
    Assert(False, 'Line1' + #13#10 + 'Line2');
  except
    on E: Exception do WriteLn(E.Message);
  end;
end.
"#);
    assert_eq!(out, vec!["Line1", "Line2"]);
}
