//! Exception hierarchy: catching base classes, custom exceptions, `AggregateException`.
use super::helpers::run_csharp;

#[test]
fn catch_base_exception_catches_derived_type() {
    assert_eq!(
        run_csharp(r#"string r="";
try{int[] a=new int[3]; var _=a[10];}
catch(System.Exception ex){r=ex.GetType().Name;}
Console.WriteLine(r);"#),
        &["IndexOutOfRangeException"]
    );
}

#[test]
fn custom_exception_stores_custom_message() {
    assert_eq!(
        run_csharp(r#"class AppEx:System.Exception{public AppEx(string m):base(m){}}
string r="";
try{throw new AppEx("fail");}
catch(AppEx ex){r=ex.Message;}
Console.WriteLine(r);"#),
        &["fail"]
    );
}

#[test]
fn custom_exception_with_inner_exception_chain() {
    assert_eq!(
        run_csharp(r#"class Outer:System.Exception{public Outer(System.Exception inner):base("outer",inner){}}
string r="";
try{throw new Outer(new System.ArgumentNullException("arg"));}
catch(Outer ex){r=ex.InnerException?.GetType().Name;}
Console.WriteLine(r);"#),
        &["ArgumentNullException"]
    );
}

#[test]
fn aggregate_exception_wraps_multiple_inner_exceptions() {
    assert_eq!(
        run_csharp(r#"var ae=new System.AggregateException(
    new System.Exception("one"),
    new System.Exception("two"));
Console.WriteLine(ae.InnerExceptions.Count);"#),
        &["2"]
    );
}

#[test]
fn exception_data_dictionary_stores_arbitrary_key_value() {
    assert_eq!(
        run_csharp(r#"string r="";
try{
    var ex=new System.Exception("test");
    ex.Data["userId"]=42;
    throw ex;
}catch(System.Exception ex){r=ex.Data["userId"].ToString();}
Console.WriteLine(r);"#),
        &["42"]
    );
}

#[test]
fn exception_source_property_set_programmatically() {
    assert_eq!(
        run_csharp(r#"var ex=new System.Exception("e");
ex.Source="MyModule";
Console.WriteLine(ex.Source);"#),
        &["MyModule"]
    );
}
