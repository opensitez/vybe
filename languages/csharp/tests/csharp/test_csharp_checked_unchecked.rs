//! `checked` and `unchecked` contexts for integer overflow.
use super::helpers::run_csharp;

#[test]
fn checked_block_throws_on_int_overflow() {
    assert_eq!(
        run_csharp(
            r#"string r="ok";
try{checked{int x=int.MaxValue;x++;}}
catch(System.OverflowException){r="overflow";}
Console.WriteLine(r);"#
        ),
        &["overflow"]
    );
}

#[test]
fn unchecked_block_wraps_silently_on_int_overflow() {
    assert_eq!(
        run_csharp(
            r#"unchecked{int x=int.MaxValue; x++; Console.WriteLine(x==int.MinValue);}
"#
        ),
        &["True"]
    );
}

#[test]
fn checked_expression_throws_on_byte_overflow() {
    assert_eq!(
        run_csharp(
            r#"string r="";
try{byte b=checked((byte)256);}
catch(System.OverflowException){r="ov";}
Console.WriteLine(r);"#
        ),
        &["ov"]
    );
}

#[test]
fn unchecked_expression_wraps_byte_overflow() {
    assert_eq!(
        run_csharp(r#"byte b=unchecked((byte)256); Console.WriteLine(b);"#),
        &["0"]
    );
}

#[test]
fn default_arithmetic_is_unchecked_for_performance() {
    assert_eq!(
        run_csharp(
            r#"int x=int.MaxValue; x++;
Console.WriteLine(x==int.MinValue);"#
        ),
        &["True"]
    );
}

#[test]
fn checked_multiply_throws_on_overflow() {
    assert_eq!(
        run_csharp(
            r#"string r="";
try{checked{int x=int.MaxValue*2;}}
catch(System.OverflowException){r="ov";}
Console.WriteLine(r);"#
        ),
        &["ov"]
    );
}
