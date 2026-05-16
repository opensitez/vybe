/// Comprehensive Pascal programs testing real-world patterns and algorithms.

use super::helpers::run_pascal;

// ===================================================================
// SORTING
// ===================================================================

#[test] fn bubble_sort() {
    assert_eq!(run_pascal(r#"program T;
var a: array of Integer; i, j, tmp: Integer;
begin
  a := [5, 3, 8, 1, 2];
  for i := 0 to High(a) - 1 do
    for j := 0 to High(a) - 1 - i do
      if a[j] > a[j + 1] then
      begin
        tmp := a[j]; a[j] := a[j + 1]; a[j + 1] := tmp;
      end;
  for i := 0 to High(a) do WriteLn(a[i]);
end."#), &["1", "2", "3", "5", "8"]);
}

#[test] fn selection_sort() {
    assert_eq!(run_pascal(r#"program T;
var a: array of Integer; i, j, minIdx, tmp: Integer;
begin
  a := [64, 25, 12, 22, 11];
  for i := 0 to High(a) - 1 do
  begin
    minIdx := i;
    for j := i + 1 to High(a) do
      if a[j] < a[minIdx] then minIdx := j;
    tmp := a[i]; a[i] := a[minIdx]; a[minIdx] := tmp;
  end;
  for i := 0 to High(a) do WriteLn(a[i]);
end."#), &["11", "12", "22", "25", "64"]);
}

// ===================================================================
// SEARCHING
// ===================================================================

#[test] fn linear_search() {
    assert_eq!(run_pascal(r#"program T;
function Find(a: array of Integer; target: Integer): Integer;
var i: Integer;
begin
  Result := -1;
  for i := 0 to High(a) do
    if a[i] = target then begin Result := i; Exit(i); end;
end;
begin
  WriteLn(Find([10, 20, 30, 40, 50], 30));
  WriteLn(Find([10, 20, 30, 40, 50], 99));
end."#), &["2", "-1"]);
}

#[test] fn binary_search() {
    assert_eq!(run_pascal(r#"program T;
function BinSearch(a: array of Integer; target: Integer): Integer;
var lo, hi, mid: Integer;
begin
  lo := 0; hi := High(a); Result := -1;
  while lo <= hi do
  begin
    mid := (lo + hi) div 2;
    if a[mid] = target then begin Result := mid; Exit(mid); end
    else if a[mid] < target then lo := mid + 1
    else hi := mid - 1;
  end;
end;
begin
  WriteLn(BinSearch([1, 3, 5, 7, 9, 11, 13], 7));
  WriteLn(BinSearch([1, 3, 5, 7, 9, 11, 13], 6));
end."#), &["3", "-1"]);
}

// ===================================================================
// STRING ALGORITHMS
// ===================================================================

#[test] fn palindrome_check() {
    assert_eq!(run_pascal(r#"program T;
function IsPalindrome(s: String): Boolean;
var i: Integer;
begin
  Result := true;
  for i := 0 to Length(s) div 2 - 1 do
    if s[i] <> s[Length(s) - 1 - i] then begin Result := false; Exit(false); end;
end;
begin
  if IsPalindrome('racecar') then WriteLn('yes') else WriteLn('no');
  if IsPalindrome('hello') then WriteLn('yes') else WriteLn('no');
end."#), &["yes", "no"]);
}

#[test] fn word_count() {
    assert_eq!(run_pascal(r#"program T;
var s: String; i, count: Integer; inWord: Boolean;
begin
  s := 'hello world how are you';
  count := 0; inWord := false;
  for i := 0 to Length(s) - 1 do
  begin
    if s[i] = ' ' then inWord := false
    else if not inWord then begin Inc(count); inWord := true; end;
  end;
  WriteLn(count);
end."#), &["5"]);
}

// ===================================================================
// MATH ALGORITHMS
// ===================================================================

#[test] fn fibonacci_iterative() {
    assert_eq!(run_pascal(r#"program T;
function Fib(n: Integer): Integer;
var a, b, i, tmp: Integer;
begin
  a := 0; b := 1;
  for i := 2 to n do begin tmp := b; b := a + b; a := tmp; end;
  if n = 0 then Result := 0 else Result := b;
end;
begin
  WriteLn(Fib(0)); WriteLn(Fib(1)); WriteLn(Fib(10));
end."#), &["0", "1", "55"]);
}

#[test] fn power_recursive() {
    assert_eq!(run_pascal(r#"program T;
function Pow(base, exp: Integer): Integer;
begin
  if exp = 0 then Result := 1
  else Result := base * Pow(base, exp - 1);
end;
begin
  WriteLn(Pow(2, 0));
  WriteLn(Pow(2, 10));
  WriteLn(Pow(3, 4));
end."#), &["1", "1024", "81"]);
}

#[test] fn lcm_gcd() {
    assert_eq!(run_pascal(r#"program T;
function GCD(a, b: Integer): Integer;
begin
  while b <> 0 do
  begin
    Result := b;
    b := a mod b;
    a := Result;
  end;
  Result := a;
end;
function LCM(a, b: Integer): Integer;
begin Result := (a * b) div GCD(a, b); end;
begin
  WriteLn(GCD(12, 8));
  WriteLn(LCM(12, 8));
end."#), &["4", "24"]);
}

// ===================================================================
// CLASS-BASED PATTERNS
// ===================================================================

#[test] fn stack_class() {
    assert_eq!(run_pascal(r#"program T;
type TStack = class
  public
    FItems: array of Integer;
    FCount: Integer;
    constructor Create;
    procedure Push(val: Integer);
    function Pop: Integer;
    function Peek: Integer;
    function IsEmpty: Boolean;
end;
constructor TStack.Create; begin FItems := []; FCount := 0; end;
procedure TStack.Push(val: Integer);
begin
  FItems := FItems;
  FCount := FCount + 1;
end;
function TStack.Pop: Integer;
begin
  FCount := FCount - 1;
  Result := FItems[FCount];
end;
function TStack.Peek: Integer;
begin Result := FItems[FCount - 1]; end;
function TStack.IsEmpty: Boolean;
begin Result := FCount = 0; end;
var s: TStack;
begin
  s := TStack.Create;
  WriteLn(BoolToStr(s.IsEmpty()));
end."#), &["true"]);
}

// ===================================================================
// COMPLEX CONTROL FLOW
// ===================================================================

#[test] fn nested_loops_with_break() {
    assert_eq!(run_pascal(r#"program T;
var i, j: Integer; found: Boolean;
begin
  found := false;
  for i := 1 to 10 do
  begin
    for j := 1 to 10 do
    begin
      if i * j = 42 then
      begin
        WriteLn('Found: ' + IntToStr(i) + ' x ' + IntToStr(j));
        found := true;
        break;
      end;
    end;
    if found then break;
  end;
end."#), &["Found: 6 x 7"]);
}

#[test] fn fizzbuzz_with_case() {
    assert_eq!(run_pascal(r#"program T;
var i: Integer;
begin
  for i := 1 to 15 do
  begin
    if (i mod 15 = 0) then WriteLn('FizzBuzz')
    else if (i mod 3 = 0) then WriteLn('Fizz')
    else if (i mod 5 = 0) then WriteLn('Buzz')
    else WriteLn(i);
  end;
end."#), &["1", "2", "Fizz", "4", "Buzz", "Fizz", "7", "8", "Fizz", "Buzz", "11", "Fizz", "13", "14", "FizzBuzz"]);
}

// ===================================================================
// TYPE CONVERSION AND MIXED OPERATIONS
// ===================================================================

#[test] fn mixed_int_real() {
    assert_eq!(run_pascal("program T; begin WriteLn(5 + 0.5); end."), &["5.5"]);
}

#[test] fn string_concatenation_chain() {
    assert_eq!(run_pascal(r#"program T;
var name: String; age: Integer;
begin
  name := 'Alice';
  age := 30;
  WriteLn('Name: ' + name + ', Age: ' + IntToStr(age));
end."#), &["Name: Alice, Age: 30"]);
}

// ===================================================================
// MULTI-LINE PROGRAMS
// ===================================================================

#[test] fn calculator_program() {
    assert_eq!(run_pascal(r#"program Calculator;
function Calc(a: Integer; op: String; b: Integer): Integer;
begin
  case op of
    '+': Result := a + b;
    '-': Result := a - b;
    '*': Result := a * b;
  else
    Result := 0;
  end;
end;
begin
  WriteLn(Calc(10, '+', 5));
  WriteLn(Calc(10, '-', 3));
  WriteLn(Calc(10, '*', 4));
end."#), &["15", "7", "40"]);
}

#[test] fn matrix_multiply_2x2() {
    assert_eq!(run_pascal(r#"program T;
var a00, a01, a10, a11: Integer;
var b00, b01, b10, b11: Integer;
var c00, c01, c10, c11: Integer;
begin
  a00 := 1; a01 := 2; a10 := 3; a11 := 4;
  b00 := 5; b01 := 6; b10 := 7; b11 := 8;
  c00 := a00*b00 + a01*b10;
  c01 := a00*b01 + a01*b11;
  c10 := a10*b00 + a11*b10;
  c11 := a10*b01 + a11*b11;
  WriteLn(c00); WriteLn(c01); WriteLn(c10); WriteLn(c11);
end."#), &["19", "22", "43", "50"]);
}
