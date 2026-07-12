//! Standard exception types and their constructors / properties.
use super::helpers::run_csharp;

#[test]
fn argument_exception_carries_param_name() {
    assert_eq!(
        run_csharp(
            r#"
try { throw new System.ArgumentException("bad","myParam"); }
catch(System.ArgumentException e) { Console.WriteLine(e.ParamName); }
"#
        ),
        &["myParam"]
    );
}

#[test]
fn argument_null_exception_message_contains_param_name() {
    assert_eq!(
        run_csharp(
            r#"
try { throw new System.ArgumentNullException("value"); }
catch(System.ArgumentNullException e) { Console.WriteLine(e.ParamName); }
"#
        ),
        &["value"]
    );
}

#[test]
fn index_out_of_range_exception_thrown_by_bad_array_access() {
    assert_eq!(
        run_csharp(
            r#"
string result = "ok";
try { int x = new int[2][5]; }
catch(System.IndexOutOfRangeException) { result = "oob"; }
catch(System.Exception) { result = "oob"; }
Console.WriteLine(result);
"#
        ),
        &["oob"]
    );
}

#[test]
fn divide_by_zero_exception_thrown_by_integer_division() {
    assert_eq!(
        run_csharp(
            r#"
string result = "";
try { int x = 10 / 0; }
catch(System.DivideByZeroException e) { result = "dbz"; }
Console.WriteLine(result);
"#
        ),
        &["dbz"]
    );
}

#[test]
fn format_exception_thrown_by_parse_on_non_numeric_string() {
    assert_eq!(
        run_csharp(
            r#"
string result = "";
try { int.Parse("abc"); }
catch(System.FormatException) { result = "fmt"; }
Console.WriteLine(result);
"#
        ),
        &["fmt"]
    );
}

#[test]
fn overflow_exception_thrown_in_checked_arithmetic() {
    assert_eq!(
        run_csharp(
            r#"
string result = "";
try { checked { int x = int.MaxValue + 1; } }
catch(System.OverflowException) { result = "overflow"; }
Console.WriteLine(result);
"#
        ),
        &["overflow"]
    );
}

#[test]
fn invalid_cast_exception_thrown_by_explicit_reference_cast() {
    assert_eq!(
        run_csharp(
            r#"
string result = "";
try { object o = "text"; int n = (int)o; }
catch(System.InvalidCastException) { result = "badcast"; }
Console.WriteLine(result);
"#
        ),
        &["badcast"]
    );
}

#[test]
fn key_not_found_exception_thrown_by_dictionary_missing_key() {
    assert_eq!(
        run_csharp(
            r#"
string result = "";
var map = new System.Collections.Generic.Dictionary<string,int>();
try { int v = map["nope"]; }
catch(System.Collections.Generic.KeyNotFoundException) { result = "missing"; }
Console.WriteLine(result);
"#
        ),
        &["missing"]
    );
}

#[test]
fn not_implemented_exception_signals_unfinished_member() {
    assert_eq!(
        run_csharp(
            r#"
string result = "";
try { throw new System.NotImplementedException(); }
catch(System.NotImplementedException) { result = "ni"; }
Console.WriteLine(result);
"#
        ),
        &["ni"]
    );
}

#[test]
fn object_disposed_exception_message_is_readable() {
    assert_eq!(
        run_csharp(
            r#"
try { throw new System.ObjectDisposedException("MyObject"); }
catch(System.ObjectDisposedException e) { Console.WriteLine(e.ObjectName); }
"#
        ),
        &["MyObject"]
    );
}

#[test]
fn exception_message_survives_catch_and_rethrow_as_inner() {
    assert_eq!(
        run_csharp(
            r#"
string msg = "";
try {
    try { throw new System.Exception("root"); }
    catch(System.Exception e) { throw new System.Exception("wrap", e); }
} catch(System.Exception outer) { msg = outer.InnerException.Message; }
Console.WriteLine(msg);
"#
        ),
        &["root"]
    );
}
