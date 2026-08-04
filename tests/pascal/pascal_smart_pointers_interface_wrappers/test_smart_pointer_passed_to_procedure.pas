// vybe-test: pascal/pascal_smart_pointers_interface_wrappers/test_smart_pointer_passed_to_procedure
// origin: languages/pascal/tests/pascal/test_pascal_smart_pointers_interface_wrappers.rs
program Test;
{$mode delphi}
uses SysUtils;
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
type TPayload = class
  public Code: Integer;
  constructor Create(C: Integer);
end;
constructor TPayload.Create(C: Integer); begin Code := C; end;

type IPayloadSmart = interface
  ['{00001111-2222-3333-4444-555566667777}']
  function GetPayload: TPayload;
end;
type TPayloadSmartImpl = class(TInterfacedObject, IPayloadSmart)
  private FPayload: TPayload;
  public constructor Create(p: TPayload); destructor Destroy; override;
  public function GetPayload: TPayload;
end;
constructor TPayloadSmartImpl.Create(p: TPayload); begin FPayload := p; end;
destructor TPayloadSmartImpl.Destroy; begin FPayload.Free; inherited Destroy; end;
function TPayloadSmartImpl.GetPayload: TPayload; begin Result := FPayload; end;

procedure ProcessSmart(smart: IPayloadSmart);
begin
  __p(__vs(smart.GetPayload.Code));
end;

var sp: IPayloadSmart;
begin
  sp := TPayloadSmartImpl.Create(TPayload.Create(999));
  ProcessSmart(sp);
__vybeCheck('999');
end.
