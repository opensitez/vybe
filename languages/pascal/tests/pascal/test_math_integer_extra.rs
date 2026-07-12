/// Gcd-style, prime checks, factorial, and small fibonacci.
use super::helpers::run_pascal;

#[test]
fn gcd_basic_twelve_eight() {
    assert_eq!(
        run_pascal(
            r#"program T; function Gcd(a,b:Integer):Integer; begin if b=0 then Result:=a else Result:=Gcd(b,a mod b); end; begin WriteLn(Gcd(12,8)); end."#
        ),
        &["4"]
    );
}

#[test]
fn gcd_coprime_seven_fifteen() {
    assert_eq!(
        run_pascal(
            r#"program T; function Gcd(a,b:Integer):Integer; begin if b=0 then Result:=a else Result:=Gcd(b,a mod b); end; begin WriteLn(Gcd(7,15)); end."#
        ),
        &["1"]
    );
}

#[test]
fn gcd_one_argument_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; function Gcd(a,b:Integer):Integer; begin if b=0 then Result:=a else Result:=Gcd(b,a mod b); end; begin WriteLn(Gcd(9,0)); end."#
        ),
        &["9"]
    );
}

#[test]
fn gcd_equal_numbers() {
    assert_eq!(
        run_pascal(
            r#"program T; function Gcd(a,b:Integer):Integer; begin if b=0 then Result:=a else Result:=Gcd(b,a mod b); end; begin WriteLn(Gcd(6,6)); end."#
        ),
        &["6"]
    );
}

#[test]
fn gcd_large_pair() {
    assert_eq!(
        run_pascal(
            r#"program T; function Gcd(a,b:Integer):Integer; begin if b=0 then Result:=a else Result:=Gcd(b,a mod b); end; begin WriteLn(Gcd(270,192)); end."#
        ),
        &["6"]
    );
}

#[test]
fn lcm_from_gcd_formula() {
    assert_eq!(
        run_pascal(
            r#"program T; function Gcd(a,b:Integer):Integer; begin if b=0 then Result:=a else Result:=Gcd(b,a mod b); end; function Lcm(a,b:Integer):Integer; begin Result:=(a div Gcd(a,b))*b; end; begin WriteLn(Lcm(4,6)); end."#
        ),
        &["12"]
    );
}

#[test]
fn is_prime_two_true() {
    assert_eq!(
        run_pascal(
            r#"program T; function IsPrime(n:Integer):Boolean; var i:Integer; begin Result:=n>1; if Result then for i:=2 to n-1 do if (n mod i)=0 then Result:=false; end; begin WriteLn(IsPrime(2)); end."#
        ),
        &["true"]
    );
}

#[test]
fn is_prime_four_false() {
    assert_eq!(
        run_pascal(
            r#"program T; function IsPrime(n:Integer):Boolean; var i:Integer; begin Result:=n>1; if Result then for i:=2 to n-1 do if (n mod i)=0 then Result:=false; end; begin WriteLn(IsPrime(4)); end."#
        ),
        &["false"]
    );
}

#[test]
fn is_prime_seventeen_true() {
    assert_eq!(
        run_pascal(
            r#"program T; function IsPrime(n:Integer):Boolean; var i:Integer; begin Result:=n>1; if Result then for i:=2 to n-1 do if (n mod i)=0 then Result:=false; end; begin WriteLn(IsPrime(17)); end."#
        ),
        &["true"]
    );
}

#[test]
fn is_prime_one_false() {
    assert_eq!(
        run_pascal(
            r#"program T; function IsPrime(n:Integer):Boolean; var i:Integer; begin Result:=n>1; if Result then for i:=2 to n-1 do if (n mod i)=0 then Result:=false; end; begin WriteLn(IsPrime(1)); end."#
        ),
        &["false"]
    );
}

#[test]
fn is_prime_nine_false() {
    assert_eq!(
        run_pascal(
            r#"program T; function IsPrime(n:Integer):Boolean; var i:Integer; begin Result:=n>1; if Result then for i:=2 to n-1 do if (n mod i)=0 then Result:=false; end; begin WriteLn(IsPrime(9)); end."#
        ),
        &["false"]
    );
}

#[test]
fn factorial_zero_one() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fact(n:Integer):Integer; begin if n<=1 then Result:=1 else Result:=n*Fact(n-1); end; begin WriteLn(Fact(0)); end."#
        ),
        &["1"]
    );
}

#[test]
fn factorial_five_one_twenty() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fact(n:Integer):Integer; begin if n<=1 then Result:=1 else Result:=n*Fact(n-1); end; begin WriteLn(Fact(5)); end."#
        ),
        &["120"]
    );
}

#[test]
fn factorial_six_seven_twenty() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fact(n:Integer):Integer; begin if n<=1 then Result:=1 else Result:=n*Fact(n-1); end; begin WriteLn(Fact(6)); end."#
        ),
        &["720"]
    );
}

#[test]
fn factorial_one_is_one() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fact(n:Integer):Integer; begin if n<=1 then Result:=1 else Result:=n*Fact(n-1); end; begin WriteLn(Fact(1)); end."#
        ),
        &["1"]
    );
}

#[test]
fn factorial_three_six() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fact(n:Integer):Integer; begin if n<=1 then Result:=1 else Result:=n*Fact(n-1); end; begin WriteLn(Fact(3)); end."#
        ),
        &["6"]
    );
}

#[test]
fn fibonacci_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fib(n:Integer):Integer; begin if n<=0 then Result:=0 else if n=1 then Result:=1 else Result:=Fib(n-1)+Fib(n-2); end; begin WriteLn(Fib(0)); end."#
        ),
        &["0"]
    );
}

#[test]
fn fibonacci_one() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fib(n:Integer):Integer; begin if n<=0 then Result:=0 else if n=1 then Result:=1 else Result:=Fib(n-1)+Fib(n-2); end; begin WriteLn(Fib(1)); end."#
        ),
        &["1"]
    );
}

#[test]
fn fibonacci_ten_fifty_five() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fib(n:Integer):Integer; begin if n<=0 then Result:=0 else if n=1 then Result:=1 else Result:=Fib(n-1)+Fib(n-2); end; begin WriteLn(Fib(10)); end."#
        ),
        &["55"]
    );
}

#[test]
fn fibonacci_six_eight() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fib(n:Integer):Integer; begin if n<=0 then Result:=0 else if n=1 then Result:=1 else Result:=Fib(n-1)+Fib(n-2); end; begin WriteLn(Fib(6)); end."#
        ),
        &["8"]
    );
}

#[test]
fn fibonacci_seven_thirteen() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fib(n:Integer):Integer; begin if n<=0 then Result:=0 else if n=1 then Result:=1 else Result:=Fib(n-1)+Fib(n-2); end; begin WriteLn(Fib(7)); end."#
        ),
        &["13"]
    );
}

#[test]
fn count_primes_up_to_ten() {
    assert_eq!(
        run_pascal(
            r#"program T; function IsPrime(n:Integer):Boolean; var i:Integer; begin Result:=n>1; if Result then for i:=2 to n-1 do if (n mod i)=0 then Result:=false; end; var n,c:Integer; begin c:=0; for n:=2 to 10 do if IsPrime(n) then Inc(c); WriteLn(c); end."#
        ),
        &["4"]
    );
}

#[test]
fn sum_factorials_one_to_four() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fact(n:Integer):Integer; begin if n<=1 then Result:=1 else Result:=n*Fact(n-1); end; var i,s:Integer; begin s:=0; for i:=1 to 4 do s:=s+Fact(i); WriteLn(s); end."#
        ),
        &["33"]
    );
}

#[test]
fn gcd_iterative_style() {
    assert_eq!(
        run_pascal(
            r#"program T; function Gcd(a,b:Integer):Integer; begin while b<>0 do begin Result:=a mod b; a:=b; b:=Result; end; Result:=a; end; begin WriteLn(Gcd(48,18)); end."#
        ),
        &["6"]
    );
}

#[test]
fn lcm_coprime_is_product() {
    assert_eq!(
        run_pascal(
            r#"program T; function Gcd(a,b:Integer):Integer; begin if b=0 then Result:=a else Result:=Gcd(b,a mod b); end; function Lcm(a,b:Integer):Integer; begin Result:=(a div Gcd(a,b))*b; end; begin WriteLn(Lcm(5,7)); end."#
        ),
        &["35"]
    );
}

#[test]
fn is_prime_eleven_true() {
    assert_eq!(
        run_pascal(
            r#"program T; function IsPrime(n:Integer):Boolean; var i:Integer; begin Result:=n>1; if Result then for i:=2 to n-1 do if (n mod i)=0 then Result:=false; end; begin WriteLn(IsPrime(11)); end."#
        ),
        &["true"]
    );
}

#[test]
fn is_prime_fifteen_false() {
    assert_eq!(
        run_pascal(
            r#"program T; function IsPrime(n:Integer):Boolean; var i:Integer; begin Result:=n>1; if Result then for i:=2 to n-1 do if (n mod i)=0 then Result:=false; end; begin WriteLn(IsPrime(15)); end."#
        ),
        &["false"]
    );
}

#[test]
fn factorial_four_twenty_four() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fact(n:Integer):Integer; begin if n<=1 then Result:=1 else Result:=n*Fact(n-1); end; begin WriteLn(Fact(4)); end."#
        ),
        &["24"]
    );
}

#[test]
fn fibonacci_eight_twenty_one() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fib(n:Integer):Integer; begin if n<=0 then Result:=0 else if n=1 then Result:=1 else Result:=Fib(n-1)+Fib(n-2); end; begin WriteLn(Fib(8)); end."#
        ),
        &["21"]
    );
}

#[test]
fn gcd_negative_handling_abs() {
    assert_eq!(
        run_pascal(
            r#"program T; function Gcd(a,b:Integer):Integer; begin a:=Abs(a); b:=Abs(b); if b=0 then Result:=a else Result:=Gcd(b,a mod b); end; begin WriteLn(Gcd(-12,8)); end."#
        ),
        &["4"]
    );
}

#[test]
fn twin_prime_check_five() {
    assert_eq!(
        run_pascal(
            r#"program T; function IsPrime(n:Integer):Boolean; var i:Integer; begin Result:=n>1; if Result then for i:=2 to n-1 do if (n mod i)=0 then Result:=false; end; begin WriteLn(IsPrime(5) and IsPrime(7)); end."#
        ),
        &["true"]
    );
}

#[test]
fn factorial_iterative_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fact(n:Integer):Integer; var i:Integer; begin Result:=1; for i:=2 to n do Result:=Result*i; end; begin WriteLn(Fact(5)); end."#
        ),
        &["120"]
    );
}

#[test]
fn fibonacci_iterative_twelve() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fib(n:Integer):Integer; var a,b,i,t:Integer; begin if n<=0 then Result:=0 else begin a:=0; b:=1; for i:=2 to n do begin t:=a+b; a:=b; b:=t; end; Result:=b; end; end; begin WriteLn(Fib(12)); end."#
        ),
        &["144"]
    );
}

#[test]
fn sum_gcd_pairs_list() {
    assert_eq!(
        run_pascal(
            r#"program T; function Gcd(a,b:Integer):Integer; begin if b=0 then Result:=a else Result:=Gcd(b,a mod b); end; begin WriteLn(Gcd(14,21)+Gcd(8,12)); end."#
        ),
        &["11"]
    );
}

#[test]
fn is_prime_thirteen_true() {
    assert_eq!(
        run_pascal(
            r#"program T; function IsPrime(n:Integer):Boolean; var i:Integer; begin Result:=n>1; if Result then for i:=2 to n-1 do if (n mod i)=0 then Result:=false; end; begin WriteLn(IsPrime(13)); end."#
        ),
        &["true"]
    );
}

#[test]
fn factorial_two_is_two() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fact(n:Integer):Integer; begin if n<=1 then Result:=1 else Result:=n*Fact(n-1); end; begin WriteLn(Fact(2)); end."#
        ),
        &["2"]
    );
}

#[test]
fn fibonacci_nine_thirty_four() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fib(n:Integer):Integer; begin if n<=0 then Result:=0 else if n=1 then Result:=1 else Result:=Fib(n-1)+Fib(n-2); end; begin WriteLn(Fib(9)); end."#
        ),
        &["34"]
    );
}

#[test]
fn gcd_three_numbers_via_fold() {
    assert_eq!(
        run_pascal(
            r#"program T; function Gcd(a,b:Integer):Integer; begin if b=0 then Result:=a else Result:=Gcd(b,a mod b); end; begin WriteLn(Gcd(Gcd(24,36),60)); end."#
        ),
        &["12"]
    );
}

#[test]
fn prime_twenty_three_true() {
    assert_eq!(
        run_pascal(
            r#"program T; function IsPrime(n:Integer):Boolean; var i:Integer; begin Result:=n>1; if Result then for i:=2 to n-1 do if (n mod i)=0 then Result:=false; end; begin WriteLn(IsPrime(23)); end."#
        ),
        &["true"]
    );
}

#[test]
fn factorial_eight40320() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fact(n:Integer):Integer; begin if n<=1 then Result:=1 else Result:=n*Fact(n-1); end; begin WriteLn(Fact(8)); end."#
        ),
        &["40320"]
    );
}

#[test]
fn fibonacci_five_five() {
    assert_eq!(
        run_pascal(
            r#"program T; function Fib(n:Integer):Integer; begin if n<=0 then Result:=0 else if n=1 then Result:=1 else Result:=Fib(n-1)+Fib(n-2); end; begin WriteLn(Fib(5)); end."#
        ),
        &["5"]
    );
}

#[test]
fn lcm_small_multiples() {
    assert_eq!(
        run_pascal(
            r#"program T; function Gcd(a,b:Integer):Integer; begin if b=0 then Result:=a else Result:=Gcd(b,a mod b); end; function Lcm(a,b:Integer):Integer; begin Result:=(a div Gcd(a,b))*b; end; begin WriteLn(Lcm(3,4)); end."#
        ),
        &["12"]
    );
}
