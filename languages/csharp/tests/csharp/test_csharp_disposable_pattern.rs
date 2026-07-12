//! `IDisposable`, `using` statement, `using` declaration, `IAsyncDisposable`.
use super::helpers::run_csharp;

#[test]
fn using_statement_calls_dispose_on_exit() {
    assert_eq!(
        run_csharp(
            r#"class Resource:System.IDisposable{
    public bool Disposed;
    public void Dispose(){Disposed=true;}
}
var r=new Resource();
using(r){}
Console.WriteLine(r.Disposed);"#
        ),
        &["True"]
    );
}

#[test]
fn using_declaration_disposes_at_end_of_block() {
    assert_eq!(
        run_csharp(
            r#"class R:System.IDisposable{public bool Gone;public void Dispose(){Gone=true;}}
R r;
{using var x=new R(); r=x;}
Console.WriteLine(r.Gone);"#
        ),
        &["True"]
    );
}

#[test]
fn using_with_exception_still_disposes() {
    assert_eq!(
        run_csharp(
            r#"class R:System.IDisposable{public bool Gone;public void Dispose(){Gone=true;}}
var r=new R();
try{using(r){throw new System.Exception();}}catch{}
Console.WriteLine(r.Gone);"#
        ),
        &["True"]
    );
}

#[test]
fn memory_stream_disposed_length_unavailable() {
    assert_eq!(
        run_csharp(
            r#"System.IO.MemoryStream ms;
using(ms=new System.IO.MemoryStream()){}
string r="";
try{var _=ms.Length;}catch(System.ObjectDisposedException){r="disposed";}
Console.WriteLine(r);"#
        ),
        &["disposed"]
    );
}

#[test]
fn try_finally_equivalent_to_using_for_cleanup() {
    assert_eq!(
        run_csharp(
            r#"bool cleaned=false;
var f=new System.Action(()=>cleaned=true);
try{}finally{f();}
Console.WriteLine(cleaned);"#
        ),
        &["True"]
    );
}
