use super::helpers::run_pascal;

#[test]
fn prog_fizzbuzz() {
    let out = run_pascal(
        r#"program T; var i: Integer;
begin for i := 1 to 15 do begin
  if (i mod 15) = 0 then WriteLn('FizzBuzz')
  else if (i mod 3) = 0 then WriteLn('Fizz')
  else if (i mod 5) = 0 then WriteLn('Buzz')
  else WriteLn(i);
end; end."#,
    );
    assert_eq!(
        out,
        &[
            "1", "2", "Fizz", "4", "Buzz", "Fizz", "7", "8", "Fizz", "Buzz", "11", "Fizz", "13",
            "14", "FizzBuzz"
        ]
    );
}

#[test]
fn prog_sum_1_to_100() {
    assert_eq!(
        run_pascal(
            "program T; var i, s: Integer; begin s := 0; for i := 1 to 100 do s := s + i; WriteLn(s); end."
        ),
        &["5050"]
    );
}

#[test]
fn prog_gcd() {
    assert_eq!(
        run_pascal(
            r#"program T;
function GCD(a, b: Integer): Integer;
begin if b = 0 then Result := a else Result := GCD(b, a mod b); end;
begin WriteLn(GCD(48, 18)); end."#
        ),
        &["6"]
    );
}

#[test]
fn prog_is_prime() {
    assert_eq!(
        run_pascal(
            r#"program T;
function IsPrime(n: Integer): Boolean;
var i: Integer;
begin
  if n < 2 then begin Result := false; Exit; end;
  i := 2;
  while i * i <= n do begin
    if (n mod i) = 0 then begin Result := false; Exit; end;
    i := i + 1;
  end;
  Result := true;
end;
begin
  WriteLn(IsPrime(2)); WriteLn(IsPrime(4)); WriteLn(IsPrime(7));
  WriteLn(IsPrime(9)); WriteLn(IsPrime(13));
end."#
        ),
        &["true", "false", "true", "false", "true"]
    );
}

#[test]
fn prog_power_iterative() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Pow(base, exp: Integer): Integer;
var i, r: Integer;
begin r := 1; for i := 1 to exp do r := r * base; Result := r; end;
begin WriteLn(Pow(2, 10)); end."#
        ),
        &["1024"]
    );
}

#[test]
fn prog_reverse_string() {
    assert_eq!(
        run_pascal(
            r#"program T;
function ReverseStr(s: String): String;
var i: Integer;
begin Result := ''; for i := Length(s) - 1 downto 0 do Result := Result + s[i]; end;
begin WriteLn(ReverseStr('hello')); end."#
        ),
        &["olleh"]
    );
}

#[test]
fn prog_count_digits() {
    assert_eq!(
        run_pascal(
            r#"program T;
function CountDigits(n: Integer): Integer;
begin
  if n < 10 then Result := 1
  else Result := 1 + CountDigits(n div 10);
end;
begin WriteLn(CountDigits(12345)); end."#
        ),
        &["5"]
    );
}

#[test]
fn prog_sum_of_digits() {
    assert_eq!(
        run_pascal(
            r#"program T;
function SumDigits(n: Integer): Integer;
begin
  if n < 10 then Result := n
  else Result := (n mod 10) + SumDigits(n div 10);
end;
begin WriteLn(SumDigits(12345)); end."#
        ),
        &["15"]
    );
}

#[test]
fn prog_max_of_three() {
    assert_eq!(
        run_pascal(
            r#"program T;
function Max3(a, b, c: Integer): Integer;
begin
  Result := a;
  if b > Result then Result := b;
  if c > Result then Result := c;
end;
begin WriteLn(Max3(3, 7, 5)); WriteLn(Max3(10, 2, 8)); end."#
        ),
        &["7", "10"]
    );
}

#[test]
fn prog_triangle_area() {
    assert_eq!(
        run_pascal(
            r#"program T;
function TriangleArea(base, height: Integer): Real;
begin Result := base * height / 2; end;
begin WriteLn(TriangleArea(10, 5)); end."#
        ),
        &["25"]
    );
}

#[test]
fn prog_accumulate_in_loop() {
    assert_eq!(
        run_pascal(
            r#"program T;
var i: Integer; var product: Integer;
begin
  product := 1;
  for i := 1 to 5 do product := product * i;
  WriteLn(product);
end."#
        ),
        &["120"]
    );
}

#[test]
fn prog_collatz_steps() {
    assert_eq!(
        run_pascal(
            r#"program T;
function CollatzSteps(n: Integer): Integer;
var steps: Integer;
begin
  steps := 0;
  while n <> 1 do begin
    if (n mod 2) = 0 then n := n div 2
    else n := 3 * n + 1;
    steps := steps + 1;
  end;
  Result := steps;
end;
begin WriteLn(CollatzSteps(6)); end."#
        ),
        &["8"]
    );
}
