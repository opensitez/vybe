// vybe-test: pascal/pascal_custom_collection_iterators/test_linked_list_enumerator
// origin: languages/pascal/tests/pascal/test_pascal_custom_collection_iterators.rs
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
type PNode = ^TNode;
     TNode = record
       Val: Integer;
       Next: PNode;
     end;

type TListWrapper = record
  Head: PNode;
end;

type TNodeEnum = record
  private FCurr: PNode;
  public function MoveNext: Boolean;
  public function GetCurrent: Integer; property Current: Integer read GetCurrent;
end;

function TNodeEnum.MoveNext: Boolean;
begin
  if FCurr <> nil then FCurr := FCurr^.Next;
  Result := FCurr <> nil;
end;
function TNodeEnum.GetCurrent: Integer; begin Result := FCurr^.Val; end;

operator Enumerator(w: TListWrapper): TNodeEnum;
var dummyHeader: TNode;
begin
  dummyHeader.Val := 0; dummyHeader.Next := w.Head;
  Result.FCurr := @dummyHeader;
end;

var n1, n2: PNode; wrap: TListWrapper; v: Integer;
begin
  New(n1); New(n2);
  n1^.Val := 100; n1^.Next := n2;
  n2^.Val := 200; n2^.Next := nil;
  wrap.Head := n1;
  for v in wrap do
    __p(__vs(v));
  Dispose(n1); Dispose(n2);
__vybeCheck('100' + #10 + '200');
end.
