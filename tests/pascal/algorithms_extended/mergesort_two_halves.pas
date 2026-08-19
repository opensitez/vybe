// vybe-test: pascal/algorithms_extended/mergesort_two_halves
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
var a: array[0..5] of Integer;
    tmp: array[0..5] of Integer;
procedure Merge(lo, mid, hi: Integer);
var i, j, k: Integer;
begin
  i := lo; j := mid + 1; k := lo;
  // `while ... do` takes ONE statement: without this begin/end the `Inc(k)`
  // sat outside the loop and every merged element overwrote tmp[k].
  while (i <= mid) and (j <= hi) do
  begin
    if a[i] <= a[j] then begin tmp[k]:=a[i]; Inc(i); end else begin tmp[k]:=a[j]; Inc(j); end;
    Inc(k);
  end;
  while i <= mid do begin tmp[k]:=a[i]; Inc(i); Inc(k); end;
  while j <= hi do begin tmp[k]:=a[j]; Inc(j); Inc(k); end;
  for i := lo to hi do a[i] := tmp[i];
end;
procedure MSort(lo, hi: Integer);
var mid: Integer;
begin
  if lo >= hi then Exit;
  mid := (lo + hi) div 2;
  MSort(lo, mid); MSort(mid + 1, hi); Merge(lo, mid, hi);
end;
var i: Integer;
begin
  a[0]:=38; a[1]:=27; a[2]:=43; a[3]:=3; a[4]:=9; a[5]:=82;
  MSort(0, 5);
  for i := 0 to 5 do __pw(__vs(IntToStr(a[i]) + ' '));
  __p(__vs(''));
__vybeCheck('3 9 27 38 43 82 ');
end.
