use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 24: Procedural Types & Method Pointer Callbacks
// ═══════════════════════════════════════════════════════════

#[test]
fn test_procedure_pointer_invocation() {
    let out = run_pascal(
        r#"
program Test;
type TNotifyProc = procedure;
procedure MyCallback;
begin
  WriteLn('CallbackExecuted');
end;
var p: TNotifyProc;
begin
  p := MyCallback;
  p();
end.
"#,
    );
    assert_eq!(out, vec!["CallbackExecuted"]);
}

#[test]
fn test_function_pointer_arguments_and_return() {
    let out = run_pascal(
        r#"
program Test;
type TMathOp = function(a, b: Integer): Integer;
function Multiply(a, b: Integer): Integer;
begin
  Result := a * b;
end;
var op: TMathOp;
begin
  op := Multiply;
  WriteLn(op(6, 7));
end.
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_assigned_check_on_procedure_pointer() {
    let out = run_pascal(
        r#"
program Test;
type TSimpleProc = procedure;
procedure Dummy; begin WriteLn('Dummy'); end;
var p: TSimpleProc;
begin
  p := nil;
  WriteLn(Assigned(p));
  p := Dummy;
  WriteLn(Assigned(p));
  p();
end.
"#,
    );
    assert_eq!(out, vec!["False", "True", "Dummy"]);
}

#[test]
fn test_method_pointer_of_object() {
    let out = run_pascal(
        r#"
program Test;
type TEvent = procedure(msg: String) of object;
type TListener = class
  public procedure OnMessage(msg: String);
end;
procedure TListener.OnMessage(msg: String);
begin
  WriteLn('LISTENED: ' + msg);
end;
var listener: TListener;
    eventHandler: TEvent;
begin
  listener := TListener.Create;
  eventHandler := listener.OnMessage;
  eventHandler('EventTriggered');
  listener.Free;
end.
"#,
    );
    assert_eq!(out, vec!["LISTENED: EventTriggered"]);
}

#[test]
fn test_procedural_parameter_higher_order_function() {
    let out = run_pascal(
        r#"
program Test;
type TFilterFunc = function(n: Integer): Boolean;
function IsEven(n: Integer): Boolean;
begin
  Result := (n mod 2) = 0;
end;
procedure FilterAndPrint(const arr: array of Integer; filter: TFilterFunc);
var i: Integer;
begin
  for i := Low(arr) to High(arr) do
    if filter(arr[i]) then WriteLn(arr[i]);
end;
begin
  FilterAndPrint([1, 2, 3, 4, 5, 6], IsEven);
end.
"#,
    );
    assert_eq!(out, vec!["2", "4", "6"]);
}

#[test]
fn test_procedural_type_array() {
    let out = run_pascal(
        r#"
program Test;
type TStepProc = procedure;
procedure Step1; begin WriteLn('Step1'); end;
procedure Step2; begin WriteLn('Step2'); end;
procedure Step3; begin WriteLn('Step3'); end;
var steps: array[1..3] of TStepProc;
    i: Integer;
begin
  steps[1] := Step1; steps[2] := Step2; steps[3] := Step3;
  for i := 1 to 3 do
    steps[i]();
end.
"#,
    );
    assert_eq!(out, vec!["Step1", "Step2", "Step3"]);
}

#[test]
fn test_procedural_variable_in_record() {
    let out = run_pascal(
        r#"
program Test;
type TAction = procedure(val: Integer);
type TCommandRec = record
  Name: String;
  Exec: TAction;
end;
procedure DoPrint(val: Integer);
begin
  WriteLn('VAL:' + val.ToString);
end;
var cmd: TCommandRec;
begin
  cmd.Name := 'PrintCmd';
  cmd.Exec := DoPrint;
  WriteLn(cmd.Name);
  cmd.Exec(99);
end.
"#,
    );
    assert_eq!(out, vec!["PrintCmd", "VAL:99"]);
}

#[test]
fn test_procedural_variable_in_class_field() {
    let out = run_pascal(
        r#"
program Test;
type TClickEvent = procedure of object;
type TButton = class
  public OnClick: TClickEvent;
  public procedure Click;
end;
type TForm = class
  public procedure ButtonClick;
end;
procedure TButton.Click;
begin
  if Assigned(OnClick) then OnClick();
end;
procedure TForm.ButtonClick;
begin
  WriteLn('ButtonClicked');
end;
var btn: TButton; frm: TForm;
begin
  btn := TButton.Create; frm := TForm.Create;
  btn.OnClick := frm.ButtonClick;
  btn.Click;
  btn.Free; frm.Free;
end.
"#,
    );
    assert_eq!(out, vec!["ButtonClicked"]);
}

#[test]
fn test_procedural_type_returning_function_pointer() {
    let out = run_pascal(
        r#"
program Test;
type TMathFunc = function(x: Integer): Integer;
function DoubleIt(x: Integer): Integer; begin Result := x * 2; end;
function TripleIt(x: Integer): Integer; begin Result := x * 3; end;
function GetMathOp(mode: Integer): TMathFunc;
begin
  if mode = 2 then Result := DoubleIt else Result := TripleIt;
end;
var op: TMathFunc;
begin
  op := GetMathOp(2);
  WriteLn(op(10));
  op := GetMathOp(3);
  WriteLn(op(10));
end.
"#,
    );
    assert_eq!(out, vec!["20", "30"]);
}

#[test]
fn test_procedural_variable_with_var_param() {
    let out = run_pascal(
        r#"
program Test;
type TMutatorProc = procedure(var x: Integer);
procedure SquareVar(var x: Integer);
begin
  x := x * x;
end;
var proc: TMutatorProc; val: Integer;
begin
  proc := SquareVar;
  val := 5;
  proc(val);
  WriteLn(val);
end.
"#,
    );
    assert_eq!(out, vec!["25"]);
}

#[test]
fn test_procedural_variable_with_out_param() {
    let out = run_pascal(
        r#"
program Test;
type TGenProc = procedure(out s: String);
procedure ProvideString(out s: String);
begin
  s := 'ProvidedData';
end;
var gen: TGenProc; text: String;
begin
  gen := ProvideString;
  gen(text);
  WriteLn(text);
end.
"#,
    );
    assert_eq!(out, vec!["ProvidedData"]);
}

#[test]
fn test_procedural_variable_string_transform() {
    let out = run_pascal(
        r#"
program Test;
type TStrMapper = function(const s: String): String;
function ToUpperMapper(const s: String): String; begin Result := UpperCase(s); end;
function ToLowerMapper(const s: String): String; begin Result := LowerCase(s); end;
var mapper: TStrMapper;
begin
  mapper := ToUpperMapper;
  WriteLn(mapper('pascal'));
  mapper := ToLowerMapper;
  WriteLn(mapper('PASCAL'));
end.
"#,
    );
    assert_eq!(out, vec!["PASCAL", "pascal"]);
}

#[test]
fn test_procedural_variable_enum_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TColor = (cRed, cGreen, cBlue);
type TColorProc = procedure(c: TColor);
procedure PrintColor(c: TColor);
begin
  WriteLn('COLOR:' + Ord(c).ToString);
end;
var p: TColorProc;
begin
  p := PrintColor;
  p(cGreen);
end.
"#,
    );
    assert_eq!(out, vec!["COLOR:1"]);
}

#[test]
fn test_procedural_variable_with_default_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TLogProc = procedure(msg: String = 'DefaultLog');
procedure DoLog(msg: String);
begin
  WriteLn(msg);
end;
var p: TLogProc;
begin
  p := DoLog;
  p();
  p('CustomLog');
end.
"#,
    );
    assert_eq!(out, vec!["DefaultLog", "CustomLog"]);
}

#[test]
fn test_method_pointer_equality_comparison() {
    let out = run_pascal(
        r#"
program Test;
type TEvent = procedure of object;
type THandler = class
  public procedure OnRun;
end;
procedure THandler.OnRun; begin end;
var h: THandler; e1, e2: TEvent;
begin
  h := THandler.Create;
  e1 := h.OnRun; e2 := h.OnRun;
  WriteLn(e1 = e2);
  h.Free;
end.
"#,
    );
    assert_eq!(out, vec!["FALSE"]);
}

#[test]
fn test_procedural_variable_subrange_parameter() {
    let out = run_pascal(
        r#"
program Test;
type TScore = 1..100;
type TScoreProc = procedure(s: TScore);
procedure HandleScore(s: TScore);
begin
  WriteLn('SCORE:' + s.ToString);
end;
var p: TScoreProc; sc: TScore;
begin
  sc := 88;
  p := HandleScore;
  p(sc);
end.
"#,
    );
    assert_eq!(out, vec!["SCORE:88"]);
}

#[test]
fn test_procedural_variable_float_computation() {
    let out = run_pascal(
        r#"
program Test;
type TFloatOp = function(r: Real): Real;
function Half(r: Real): Real; begin Result := r / 2.0; end;
var op: TFloatOp;
begin
  op := Half;
  WriteLn(op(10.0));
end.
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn test_procedural_type_closure_like_pass() {
    let out = run_pascal(
        r#"
program Test;
type TReducer = function(acc, val: Integer): Integer;
function SumReducer(acc, val: Integer): Integer; begin Result := acc + val; end;
function Fold(const arr: array of Integer; initial: Integer; r: TReducer): Integer;
var i: Integer;
begin
  Result := initial;
  for i := Low(arr) to High(arr) do
    Result := r(Result, arr[i]);
end;
begin
  WriteLn(Fold([10, 20, 30], 0, SumReducer));
end.
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn test_method_pointer_cleared_to_nil() {
    let out = run_pascal(
        r#"
program Test;
type TEvent = procedure of object;
type TItem = class procedure Action; end;
procedure TItem.Action; begin end;
var item: TItem; ev: TEvent;
begin
  item := TItem.Create;
  ev := item.Action;
  WriteLn(Assigned(ev));
  ev := nil;
  WriteLn(Assigned(ev));
  item.Free;
end.
"#,
    );
    assert_eq!(out, vec!["TRUE", "FALSE"]);
}

#[test]
fn test_procedural_variable_boolean_predicates() {
    let out = run_pascal(
        r#"
program Test;
type TPredicate = function(s: String): Boolean;
function IsNotEmpty(s: String): Boolean; begin Result := Length(s) > 0; end;
var pred: TPredicate;
begin
  pred := IsNotEmpty;
  WriteLn(pred('hello'));
  WriteLn(pred(''));
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}
