use super::helpers::run_pascal;

// -- I/O --
#[test]
fn io_writeln_one() {
    assert_eq!(
        run_pascal("program T; begin WriteLn('hello'); end."),
        &["hello"]
    );
}
#[test]
fn io_writeln_multi() {
    assert_eq!(
        run_pascal("program T; begin WriteLn('a', 'b', 'c'); end."),
        &["a b c"]
    );
}
#[test]
fn io_writeln_int() {
    assert_eq!(run_pascal("program T; begin WriteLn(42); end."), &["42"]);
}
#[test]
fn io_writeln_real() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(3.14); end."),
        &["3.14"]
    );
}
#[test]
fn io_writeln_bool() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(true); end."),
        &["true"]
    );
}
#[test]
fn io_writeln_expr() {
    assert_eq!(run_pascal("program T; begin WriteLn(2 + 3); end."), &["5"]);
}
#[test]
fn io_multi_writeln() {
    assert_eq!(
        run_pascal("program T; begin WriteLn('a'); WriteLn('b'); WriteLn('c'); end."),
        &["a", "b", "c"]
    );
}

// -- String builtins --
#[test]
fn str_length() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Length('hello')); end."),
        &["5"]
    );
}
#[test]
fn str_length_empty() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Length('')); end."),
        &["0"]
    );
}
#[test]
fn str_uppercase() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(UpperCase('hello')); end."),
        &["HELLO"]
    );
}
#[test]
fn str_lowercase() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(LowerCase('HELLO')); end."),
        &["hello"]
    );
}
#[test]
fn str_trim() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Trim('  hi  ')); end."),
        &["hi"]
    );
}
#[test]
fn str_concat_plus() {
    assert_eq!(
        run_pascal("program T; begin WriteLn('foo' + 'bar'); end."),
        &["foobar"]
    );
}
#[test]
fn str_concat_fn() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Concat('a', 'b', 'c')); end."),
        &["abc"]
    );
}
#[test]
fn str_concat_var() {
    assert_eq!(
        run_pascal(
            "program T; var a, b: String; begin a := 'hello'; b := ' world'; WriteLn(a + b); end."
        ),
        &["hello world"]
    );
}
#[test]
fn str_multi_concat() {
    assert_eq!(
        run_pascal("program T; begin WriteLn('a' + 'b' + 'c' + 'd'); end."),
        &["abcd"]
    );
}
#[test]
fn str_inttostr() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(IntToStr(42)); end."),
        &["42"]
    );
}
#[test]
fn str_floattostr() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(FloatToStr(3.14)); end."),
        &["3.14"]
    );
}
#[test]
fn str_strtoint() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(StrToInt('42')); end."),
        &["42"]
    );
}
#[test]
fn str_strtofloat() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(StrToFloat('3.14')); end."),
        &["3.14"]
    );
}

// -- Math builtins --
#[test]
fn math_abs_pos() {
    assert_eq!(run_pascal("program T; begin WriteLn(Abs(5)); end."), &["5"]);
}
#[test]
fn math_abs_neg() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Abs(-5)); end."),
        &["5"]
    );
}
#[test]
fn math_abs_zero() {
    assert_eq!(run_pascal("program T; begin WriteLn(Abs(0)); end."), &["0"]);
}
#[test]
fn math_sqr() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Sqr(4)); end."),
        &["16"]
    );
}
#[test]
fn math_sqr_neg() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Sqr(-3)); end."),
        &["9"]
    );
}
#[test]
fn math_min() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Min(3, 7)); end."),
        &["3"]
    );
}
#[test]
fn math_max() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Max(3, 7)); end."),
        &["7"]
    );
}
#[test]
fn math_min_eq() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Min(5, 5)); end."),
        &["5"]
    );
}
#[test]
fn math_floor() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Floor(3.7)); end."),
        &["3"]
    );
}
#[test]
fn math_ceil() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Ceil(3.2)); end."),
        &["4"]
    );
}
#[test]
fn math_round() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Round(3.5)); end."),
        &["4"]
    );
}
#[test]
fn math_trunc() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Trunc(3.9)); end."),
        &["3"]
    );
}
#[test]
fn math_power() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Power(2, 10)); end."),
        &["1024"]
    );
}
#[test]
fn math_succ() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Succ(5)); end."),
        &["6"]
    );
}
#[test]
fn math_pred() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(Pred(5)); end."),
        &["4"]
    );
}

// -- Inc / Dec --
#[test]
fn inc_basic() {
    assert_eq!(
        run_pascal("program T; var x: Integer; begin x := 5; Inc(x); WriteLn(x); end."),
        &["6"]
    );
}
#[test]
fn dec_basic() {
    assert_eq!(
        run_pascal("program T; var x: Integer; begin x := 5; Dec(x); WriteLn(x); end."),
        &["4"]
    );
}
#[test]
fn inc_multiple() {
    assert_eq!(
        run_pascal(
            "program T; var x: Integer; begin x := 0; Inc(x); Inc(x); Inc(x); WriteLn(x); end."
        ),
        &["3"]
    );
}
#[test]
fn dec_to_neg() {
    assert_eq!(
        run_pascal("program T; var x: Integer; begin x := 1; Dec(x); Dec(x); WriteLn(x); end."),
        &["-1"]
    );
}
#[test]
fn inc_then_dec() {
    assert_eq!(
        run_pascal("program T; var x: Integer; begin x := 10; Inc(x); Dec(x); WriteLn(x); end."),
        &["10"]
    );
}

// -- Arrays --
#[test]
fn array_literal() {
    assert_eq!(
        run_pascal(
            "program T; var a: array of Integer; begin a := [10, 20, 30]; WriteLn(a[0]); WriteLn(a[2]); end."
        ),
        &["10", "30"]
    );
}
#[test]
fn array_assign() {
    assert_eq!(
        run_pascal(
            "program T; var a: array of Integer; begin a := [1, 2, 3]; a[1] := 99; WriteLn(a[1]); end."
        ),
        &["99"]
    );
}
#[test]
fn array_length() {
    assert_eq!(
        run_pascal(
            "program T; var a: array of Integer; begin a := [10, 20, 30]; WriteLn(Length(a)); end."
        ),
        &["3"]
    );
}
#[test]
fn array_high() {
    assert_eq!(
        run_pascal(
            "program T; var a: array of Integer; begin a := [10, 20, 30]; WriteLn(High(a)); end."
        ),
        &["2"]
    );
}
#[test]
fn array_low() {
    assert_eq!(
        run_pascal(
            "program T; var a: array of Integer; begin a := [10, 20, 30]; WriteLn(Low(a)); end."
        ),
        &["0"]
    );
}
#[test]
fn array_iterate() {
    assert_eq!(
        run_pascal(
            "program T; var a: array of Integer; var i: Integer; begin a := [5, 10, 15]; for i := 0 to High(a) do WriteLn(a[i]); end."
        ),
        &["5", "10", "15"]
    );
}

// -- Misc --
#[test]
fn assigned_nil() {
    assert_eq!(
        run_pascal("program T; begin if not Assigned(nil) then WriteLn('y'); end."),
        &["y"]
    );
}
