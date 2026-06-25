//! Nullable reference types (NRT): null-forgiving `!`, `?` annotations, runtime null checks.
use super::helpers::run_csharp;

#[test]
fn null_forgiving_operator_suppresses_warning_still_nulls_at_runtime() {
    assert_eq!(
        run_csharp(r#"string? s=null;
string r="ok";
try{Console.WriteLine(s!.Length);}
catch(System.NullReferenceException){r="null";}
Console.WriteLine(r);"#),
        &["null"]
    );
}

#[test]
fn nullable_string_can_hold_null_value() {
    assert_eq!(
        run_csharp(r#"string? s=null;
Console.WriteLine(s==null);"#),
        &["True"]
    );
}

#[test]
fn is_not_null_pattern_guards_against_null_dereference() {
    assert_eq!(
        run_csharp(r#"string? s="hello";
if(s is not null) Console.WriteLine(s.Length);"#),
        &["5"]
    );
}

#[test]
fn nullable_reference_getvalueorlength_via_null_coalescing() {
    assert_eq!(
        run_csharp(r#"string? s=null;
int len=s?.Length??-1;
Console.WriteLine(len);"#),
        &["-1"]
    );
}

#[test]
fn nullable_reference_type_array_element_can_be_null() {
    assert_eq!(
        run_csharp(r#"string?[] arr=new string?[3];
arr[0]="a"; arr[1]=null; arr[2]="c";
int nonNull=0;
foreach(var s in arr) if(s!=null) nonNull++;
Console.WriteLine(nonNull);"#),
        &["2"]
    );
}
