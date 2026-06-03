use super::helpers::run_pascal;

#[test]
fn lit_integer() {
    assert_eq!(run_pascal("program T; begin WriteLn(42); end."), &["42"]);
}
#[test]
fn lit_negative() {
    assert_eq!(run_pascal("program T; begin WriteLn(-7); end."), &["-7"]);
}
#[test]
fn lit_zero() {
    assert_eq!(run_pascal("program T; begin WriteLn(0); end."), &["0"]);
}
#[test]
fn lit_real() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(3.14); end."),
        &["3.14"]
    );
}
#[test]
fn lit_string() {
    assert_eq!(
        run_pascal("program T; begin WriteLn('hello'); end."),
        &["hello"]
    );
}
#[test]
fn lit_empty_string() {
    assert_eq!(run_pascal("program T; begin WriteLn(''); end."), &[""]);
}
#[test]
fn lit_bool_true() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(true); end."),
        &["true"]
    );
}
#[test]
fn lit_bool_false() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(false); end."),
        &["false"]
    );
}
#[test]
fn lit_nil() {
    assert_eq!(run_pascal("program T; begin WriteLn(nil); end."), &["null"]);
}
#[test]
fn lit_large_int() {
    assert_eq!(
        run_pascal("program T; begin WriteLn(1000000); end."),
        &["1000000"]
    );
}
#[test]
fn lit_string_spaces() {
    assert_eq!(
        run_pascal("program T; begin WriteLn('  hi  '); end."),
        &["  hi  "]
    );
}
