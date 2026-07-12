use super::helpers::run_pascal;

#[test]
fn test_programs5_caesar_encrypt() {
    let src = r#"
program T;
function Caesar(s: string; shift: Integer): string;
var
  i: Integer;
  ch: Char;
begin
  Result := '';
  for i := 1 to Length(s) do begin
    ch := s[i];
    if (ch >= 'a') and (ch <= 'z') then
      Result := Result + Chr((Ord(ch) - Ord('a') + shift) mod 26 + Ord('a'))
    else if (ch >= 'A') and (ch <= 'Z') then
      Result := Result + Chr((Ord(ch) - Ord('A') + shift) mod 26 + Ord('A'))
    else
      Result := Result + ch;
  end;
end;
begin
  WriteLn(Caesar('Hello', 3));
  WriteLn(Caesar('Khoor', 23));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["Khoor", "Hello"]);
}

#[test]
fn test_programs5_mean_calc() {
    let src = r#"
program T;
var
  data: array[1..5] of Integer;
  sum, i: Integer;
  mean: Double;
begin
  data[1] := 10;
  data[2] := 20;
  data[3] := 30;
  data[4] := 40;
  data[5] := 50;
  sum := 0;
  for i := 1 to 5 do
    sum := sum + data[i];
  mean := sum / 5.0;
  WriteLn(sum);
  WriteLn(mean);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["150", "30"]);
}

#[test]
fn test_programs5_luhn_check() {
    let src = r#"
program T;
function LuhnCheck(s: string): Boolean;
var
  i, n, digit, sum: Integer;
  double: Boolean;
begin
  sum := 0;
  double := false;
  for i := Length(s) downto 1 do begin
    digit := Ord(s[i]) - Ord('0');
    if double then begin
      digit := digit * 2;
      if digit > 9 then digit := digit - 9;
    end;
    sum := sum + digit;
    double := not double;
  end;
  Result := (sum mod 10) = 0;
end;
begin
  if LuhnCheck('4532015112830366') then WriteLn('valid') else WriteLn('invalid');
  if LuhnCheck('1234567890123456') then WriteLn('valid') else WriteLn('invalid');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["valid", "invalid"]);
}

#[test]
fn test_programs5_base_convert() {
    let src = r#"
program T;
function IntToBin(n: Integer): string;
begin
  Result := '';
  if n = 0 then begin
    Result := '0';
    Exit;
  end;
  while n > 0 do begin
    if (n mod 2) = 1 then
      Result := '1' + Result
    else
      Result := '0' + Result;
    n := n div 2;
  end;
end;
begin
  WriteLn(IntToBin(0));
  WriteLn(IntToBin(5));
  WriteLn(IntToBin(10));
  WriteLn(IntToBin(255));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["0", "101", "1010", "11111111"]);
}

#[test]
fn test_programs5_run_length_encode() {
    let src = r#"
program T;
function RLE(s: string): string;
var
  i, count: Integer;
  ch: Char;
begin
  Result := '';
  if Length(s) = 0 then Exit;
  ch := s[1];
  count := 1;
  for i := 2 to Length(s) do begin
    if s[i] = ch then
      count := count + 1
    else begin
      Result := Result + IntToStr(count) + ch;
      ch := s[i];
      count := 1;
    end;
  end;
  Result := Result + IntToStr(count) + ch;
end;
begin
  WriteLn(RLE('aabbbcccc'));
  WriteLn(RLE('abcd'));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["2a3b4c", "1a1b1c1d"]);
}

#[test]
fn test_programs5_word_count() {
    let src = r#"
program T;
function WordCount(s: string): Integer;
var
  i: Integer;
  inWord: Boolean;
begin
  Result := 0;
  inWord := false;
  for i := 1 to Length(s) do begin
    if s[i] <> ' ' then begin
      if not inWord then begin
        Result := Result + 1;
        inWord := true;
      end;
    end else
      inWord := false;
  end;
end;
begin
  WriteLn(WordCount('hello world'));
  WriteLn(WordCount('one two three four'));
  WriteLn(WordCount('  leading spaces'));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["2", "4", "2"]);
}

#[test]
fn test_programs5_gcd_lcm() {
    let src = r#"
program T;
function GCD(a, b: Integer): Integer;
begin
  while b <> 0 do begin
    a := a mod b;
    if a = 0 then begin
      a := b;
      b := 0;
    end else begin
      b := a mod b;
      a := a - a mod b;
    end;
  end;
  Result := a;
end;

function GCD2(a, b: Integer): Integer;
begin
  if b = 0 then Result := a
  else Result := GCD2(b, a mod b);
end;

begin
  WriteLn(GCD2(48, 18));
  WriteLn(GCD2(100, 75));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["6", "25"]);
}

#[test]
fn test_programs5_prime_count() {
    let src = r#"
program T;
function IsPrime(n: Integer): Boolean;
var
  i: Integer;
begin
  if n < 2 then begin Result := false; Exit; end;
  if n = 2 then begin Result := true; Exit; end;
  if n mod 2 = 0 then begin Result := false; Exit; end;
  i := 3;
  while i * i <= n do begin
    if n mod i = 0 then begin Result := false; Exit; end;
    i := i + 2;
  end;
  Result := true;
end;
var
  i, count: Integer;
begin
  count := 0;
  for i := 2 to 30 do
    if IsPrime(i) then count := count + 1;
  WriteLn(count);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_programs5_matrix_trace() {
    let src = r#"
program T;
var
  m: array[1..3, 1..3] of Integer;
  trace, i: Integer;
begin
  m[1,1] := 1; m[1,2] := 2; m[1,3] := 3;
  m[2,1] := 4; m[2,2] := 5; m[2,3] := 6;
  m[3,1] := 7; m[3,2] := 8; m[3,3] := 9;
  trace := 0;
  for i := 1 to 3 do
    trace := trace + m[i, i];
  WriteLn(trace);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_programs5_reverse_string() {
    let src = r#"
program T;
function Reverse(s: string): string;
var
  i: Integer;
begin
  Result := '';
  for i := Length(s) downto 1 do
    Result := Result + s[i];
end;
begin
  WriteLn(Reverse('hello'));
  WriteLn(Reverse('abcde'));
  WriteLn(Reverse(''));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["olleh", "edcba", ""]);
}

#[test]
fn test_programs5_fibonacci_memo() {
    let src = r#"
program T;
var
  memo: array[0..20] of Integer;

function Fib(n: Integer): Integer;
begin
  if memo[n] > 0 then begin
    Result := memo[n];
    Exit;
  end;
  if n <= 1 then
    Result := n
  else
    Result := Fib(n - 1) + Fib(n - 2);
  memo[n] := Result;
end;

var
  i: Integer;
begin
  for i := 0 to 20 do
    memo[i] := 0;
  for i := 0 to 7 do
    Write(IntToStr(Fib(i)) + ' ');
  WriteLn('');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["0 1 1 2 3 5 8 13 "]);
}

#[test]
fn test_programs5_histogram() {
    let src = r#"
program T;
var
  freq: array['a'..'e'] of Integer;
  s: string;
  i: Integer;
  ch: Char;
begin
  s := 'abcaabba';
  for ch := 'a' to 'e' do
    freq[ch] := 0;
  for i := 1 to Length(s) do begin
    ch := s[i];
    if (ch >= 'a') and (ch <= 'e') then
      freq[ch] := freq[ch] + 1;
  end;
  WriteLn(freq['a']);
  WriteLn(freq['b']);
  WriteLn(freq['c']);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["4", "3", "1"]);
}

#[test]
fn test_programs5_bubble_sort() {
    let src = r#"
program T;
var
  arr: array[1..5] of Integer;
  i, j, tmp: Integer;
begin
  arr[1] := 5; arr[2] := 3; arr[3] := 1; arr[4] := 4; arr[5] := 2;
  for i := 1 to 4 do
    for j := 1 to 5 - i do
      if arr[j] > arr[j+1] then begin
        tmp := arr[j];
        arr[j] := arr[j+1];
        arr[j+1] := tmp;
      end;
  for i := 1 to 5 do
    Write(IntToStr(arr[i]) + ' ');
  WriteLn('');
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["1 2 3 4 5 "]);
}

#[test]
fn test_programs5_stack_calculator() {
    let src = r#"
program T;
var
  stack: array[0..9] of Integer;
  top: Integer;

procedure Push(v: Integer);
begin
  stack[top] := v;
  top := top + 1;
end;

function Pop: Integer;
begin
  top := top - 1;
  Result := stack[top];
end;

var
  a, b: Integer;
begin
  top := 0;
  Push(3);
  Push(4);
  a := Pop;
  b := Pop;
  Push(a + b);
  Push(2);
  a := Pop;
  b := Pop;
  Push(a * b);
  WriteLn(Pop);
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["14"]);
}

#[test]
fn test_programs5_anagram_check() {
    let src = r#"
program T;
function CountChar(s: string; ch: Char): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 1 to Length(s) do
    if s[i] = ch then Result := Result + 1;
end;

function IsAnagram(a, b: string): Boolean;
var
  i: Integer;
begin
  if Length(a) <> Length(b) then begin Result := false; Exit; end;
  Result := true;
  for i := 1 to Length(a) do
    if CountChar(a, a[i]) <> CountChar(b, a[i]) then begin
      Result := false;
      Exit;
    end;
end;
begin
  WriteLn(IsAnagram('listen', 'silent'));
  WriteLn(IsAnagram('hello', 'world'));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn test_programs5_binary_search() {
    let src = r#"
program T;
function BinSearch(arr: array of Integer; n, target: Integer): Integer;
var
  lo, hi, mid: Integer;
begin
  lo := 0;
  hi := n - 1;
  Result := -1;
  while lo <= hi do begin
    mid := (lo + hi) div 2;
    if arr[mid] = target then begin
      Result := mid;
      Exit;
    end else if arr[mid] < target then
      lo := mid + 1
    else
      hi := mid - 1;
  end;
end;
var
  data: array[0..4] of Integer;
begin
  data[0] := 2; data[1] := 5; data[2] := 8; data[3] := 12; data[4] := 20;
  WriteLn(BinSearch(data, 5, 8));
  WriteLn(BinSearch(data, 5, 7));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["2", "-1"]);
}

#[test]
fn test_programs5_palindrome_check() {
    let src = r#"
program T;
function IsPalindrome(s: string): Boolean;
var
  i, n: Integer;
begin
  n := Length(s);
  Result := true;
  for i := 1 to n div 2 do
    if s[i] <> s[n - i + 1] then begin
      Result := false;
      Exit;
    end;
end;
begin
  WriteLn(IsPalindrome('racecar'));
  WriteLn(IsPalindrome('hello'));
  WriteLn(IsPalindrome('level'));
  WriteLn(IsPalindrome('a'));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["true", "false", "true", "true"]);
}

#[test]
fn test_programs5_number_to_words() {
    let src = r#"
program T;
function Units(n: Integer): string;
begin
  case n of
    0: Result := 'zero';
    1: Result := 'one';
    2: Result := 'two';
    3: Result := 'three';
    4: Result := 'four';
    5: Result := 'five';
    else Result := 'many';
  end;
end;
var
  i: Integer;
begin
  for i := 0 to 5 do
    WriteLn(Units(i));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["zero", "one", "two", "three", "four", "five"]);
}

#[test]
fn test_programs5_collatz() {
    let src = r#"
program T;
function CollatzLen(n: Integer): Integer;
var
  steps: Integer;
begin
  steps := 0;
  while n <> 1 do begin
    if n mod 2 = 0 then
      n := n div 2
    else
      n := n * 3 + 1;
    steps := steps + 1;
  end;
  Result := steps;
end;
begin
  WriteLn(CollatzLen(6));
  WriteLn(CollatzLen(27));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["8", "111"]);
}

#[test]
fn test_programs5_digit_sum() {
    let src = r#"
program T;
function DigitSum(n: Integer): Integer;
begin
  Result := 0;
  while n > 0 do begin
    Result := Result + (n mod 10);
    n := n div 10;
  end;
end;
begin
  WriteLn(DigitSum(123));
  WriteLn(DigitSum(9999));
  WriteLn(DigitSum(100));
end.
"#;
    let out = run_pascal(src);
    assert_eq!(out, vec!["6", "36", "1"]);
}
