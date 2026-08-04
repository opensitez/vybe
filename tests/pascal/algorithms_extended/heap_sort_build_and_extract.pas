// vybe-test: pascal/algorithms_extended/heap_sort_build_and_extract
// origin: languages/pascal/tests/pascal/test_algorithms_extended.rs
program T;
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
var a: array[0..4] of Integer;
procedure Heapify(n, i: Integer);
var largest, l, r, tmp: Integer;
begin
  largest := i; l := 2*i+1; r := 2*i+2;
  if (l < n) and (a[l] > a[largest]) then largest := l;
  if (r < n) and (a[r] > a[largest]) then largest := r;
  if largest <> i then begin tmp:=a[i]; a[i]:=a[largest]; a[largest]:=tmp; Heapify(n, largest); end;
end;
var i, n, tmp: Integer;
begin
  a[0]:=4; a[1]:=10; a[2]:=3; a[3]:=5; a[4]:=1;
  n := 5;
  for i := n div 2 downto 0 do Heapify(n, i);
  for i := n - 1 downto 1 do
  begin tmp:=a[0]; a[0]:=a[i]; a[i]:=tmp; Heapify(i, 0); end;
  for i := 0 to 4 do __pw(__vs(IntToStr(a[i]) + ' '));
  __p(__vs(''));
__vybeCheck('1 3 4 5 10 ');
end.
