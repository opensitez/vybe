// vybe-test: pascal/pascal_comprehensive_integration_edge_cases/test_integration_generic_pipeline_with_interfaces
// origin: languages/pascal/tests/pascal/test_pascal_comprehensive_integration_edge_cases.rs
program Test;
{$mode delphi}
// Vybe test harness — Pascal.
//
// Real Pascal: this compiles with `fpc` on its own, which is what lets an
// extracted test be compared against Free Pascal.
//
// Output is COLLECTED, not paired: the emitter rewrites `WriteLn(a, b, c)`
// into `__p(__vs(a) + __vs(b) + __vs(c))` and compares the whole buffer once,
// so a program whose print count is not static — a loop — still asserts.
//
// WHY THE OVERLOAD SET.
// `WriteLn` takes any type, and the corpus calls it with string literals
// (2,011), bare typed identifiers (1,361), multi-argument lists (830) and
// arbitrary expressions (4,608). There is no single conversion that covers
// those, and Vybe has no `WriteStr` (measured: "undefined is not callable"),
// which is the one primitive that would render an argument list exactly as
// `Write` does. Overloading `__vs` lets the COMPILER pick the conversion per
// argument, which is type-agnostic without needing to know any types here.
//
// `{$mode delphi}` and `uses SysUtils` are required by fpc for `Format` and
// `IntToStr`; Vybe accepts both and ignores what it does not need. The emitter
// adds them when the source lacks them, so one file runs on both.
//
// Known divergence, NOT normalised here: `WriteLn(aBoolean)` prints `TRUE`
// under fpc and `True` under Vybe. `__vs` returns `True`/`False` to match the
// corpus, which recorded Vybe's output.

var
  __vybeOut: string;

function __vs(v: string): string; overload; begin Result := v; end;
function __vs(v: integer): string; overload; begin Result := IntToStr(v); end;
function __vs(v: int64): string; overload; begin Result := IntToStr(v); end;
function __vs(v: real): string; overload; begin Str(v, Result); Result := Trim(Result); end;
function __vs(v: boolean): string; overload;
begin
  if v then Result := 'True' else Result := 'False';
end;
function __vs(v: char): string; overload; begin Result := v; end;

// `WriteLn` ends the line; `Write` does not.
procedure __p(s: string); begin __vybeOut := __vybeOut + s + #10; end;
procedure __pw(s: string); begin __vybeOut := __vybeOut + s; end;

procedure __vybeCheck(want: string);
var got: string;
begin
  got := __vybeOut;
  // The final WriteLn contributes a trailing newline the expected line vector
  // never carried.
  if (Length(got) > 0) and (got[Length(got)] = #10) then
    got := Copy(got, 1, Length(got) - 1);
  if got <> want then
  begin
    // Printed BEFORE halting: an uncaught error renders as
    // `RuntimeError: [object]` under Vybe, losing both values.
    WriteLn('FAIL: want [', want, '] got [', got, ']');
    Halt(1);
  end;
end;
uses Generics.Collections;

type ITask = interface
  ['{11111111-1111-1111-1111-111111111111}']
  function Execute: String;
end;

type TTaskProcessor<T: ITask> = class
  private FTasks: TList<T>;
  public
    constructor Create;
    destructor Destroy; override;
    procedure AddTask(task: T);
    procedure RunAll;
end;

constructor TTaskProcessor<T>.Create; begin FTasks := TList<T>.Create; end;
destructor TTaskProcessor<T>.Destroy; begin FTasks.Free; inherited Destroy; end;
procedure TTaskProcessor<T>.AddTask(task: T); begin FTasks.Add(task); end;
procedure TTaskProcessor<T>.RunAll;
var t: T;
begin
  for t in FTasks do __p(__vs(t.Execute));
end;

type TConcreteTask = class(TInterfacedObject, ITask)
  private FName: String;
  public
    constructor Create(const N: String);
    function Execute: String;
end;
constructor TConcreteTask.Create(const N: String); begin FName := N; end;
function TConcreteTask.Execute: String; begin Result := 'TaskDone:' + FName; end;

var proc: TTaskProcessor<ITask>;
begin
  proc := TTaskProcessor<ITask>.Create;
  proc.AddTask(TConcreteTask.Create('Alpha'));
  proc.AddTask(TConcreteTask.Create('Beta'));
  proc.RunAll;
  proc.Free;
__vybeCheck('TaskDone:Alpha' + #10 + 'TaskDone:Beta');
end.
