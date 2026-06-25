//! Value type vs reference type semantics: copy vs alias, struct vs class behaviour.
use super::helpers::run_csharp;

#[test]
fn struct_assignment_copies_all_fields() {
    assert_eq!(
        run_csharp(r#"struct Pt{public int X,Y;}
var a=new Pt{X=1,Y=2};
var b=a; b.X=99;
Console.WriteLine(a.X);"#),
        &["1"]
    );
}

#[test]
fn class_assignment_creates_alias_not_copy() {
    assert_eq!(
        run_csharp(r#"class Pt{public int X,Y;}
var a=new Pt{X=1,Y=2};
var b=a; b.X=99;
Console.WriteLine(a.X);"#),
        &["99"]
    );
}

#[test]
fn boxing_wraps_value_type_in_object() {
    assert_eq!(
        run_csharp(r#"int n=42; object o=n;
Console.WriteLine(o); Console.WriteLine(o is int);"#),
        &["42", "True"]
    );
}

#[test]
fn unboxing_extracts_original_value() {
    assert_eq!(
        run_csharp(r#"object o=42; int n=(int)o;
Console.WriteLine(n);"#),
        &["42"]
    );
}

#[test]
fn passing_struct_by_value_does_not_mutate_caller() {
    assert_eq!(
        run_csharp(r#"struct S{public int V;}
void Mutate(S s){s.V=999;}
var s=new S{V=1};
Mutate(s);
Console.WriteLine(s.V);"#),
        &["1"]
    );
}

#[test]
fn passing_class_by_reference_mutates_caller() {
    assert_eq!(
        run_csharp(r#"class C{public int V;}
void Mutate(C c){c.V=999;}
var c=new C{V=1};
Mutate(c);
Console.WriteLine(c.V);"#),
        &["999"]
    );
}

#[test]
fn null_reference_throws_null_reference_exception() {
    assert_eq!(
        run_csharp(r#"string r="";
try{string s=null;int len=s.Length;}
catch(System.NullReferenceException){r="null";}
Console.WriteLine(r);"#),
        &["null"]
    );
}
