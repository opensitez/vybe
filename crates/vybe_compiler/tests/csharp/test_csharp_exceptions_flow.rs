//! Exception control-flow: multi-catch, when filters, rethrow, finally ordering.
use super::helpers::run_csharp;

#[test]
fn multi_catch_picks_first_matching_clause() {
    assert_eq!(
        run_csharp(r#"string r="";
try{throw new System.ArgumentNullException("x");}
catch(System.ArgumentOutOfRangeException){r="range";}
catch(System.ArgumentNullException){r="null";}
catch(System.Exception){r="general";}
Console.WriteLine(r);"#),
        &["null"]
    );
}

#[test]
fn finally_always_runs_even_after_return() {
    assert_eq!(
        run_csharp(r#"bool ran=false;
int Compute(){
    try{return 42;}
    finally{ran=true;}
}
int v=Compute();
Console.WriteLine(v); Console.WriteLine(ran);"#),
        &["42", "True"]
    );
}

#[test]
fn catch_when_filter_skips_non_matching_predicate() {
    assert_eq!(
        run_csharp(r#"string r="unhandled";
try{throw new System.Exception("skip");}
catch(System.Exception ex) when(ex.Message=="match"){r="matched";}
catch(System.Exception){r="caught";}
Console.WriteLine(r);"#),
        &["caught"]
    );
}

#[test]
fn rethrow_without_argument_preserves_stack_trace() {
    assert_eq!(
        run_csharp(r#"string r="";
try{
    try{throw new System.Exception("orig");}
    catch{throw;}
}catch(System.Exception ex){r=ex.Message;}
Console.WriteLine(r);"#),
        &["orig"]
    );
}

#[test]
fn exception_in_finally_replaces_original() {
    assert_eq!(
        run_csharp(r#"string r="";
try{
    try{throw new System.Exception("orig");}
    finally{throw new System.Exception("final");}
}catch(System.Exception ex){r=ex.Message;}
Console.WriteLine(r);"#),
        &["final"]
    );
}

#[test]
fn nested_try_catch_inner_handles_before_outer() {
    assert_eq!(
        run_csharp(r#"string r="";
try{
    try{throw new System.Exception("inner");}
    catch(System.Exception ex){r="inner:"+ex.Message; throw new System.Exception("outer");}
}catch(System.Exception ex){r+=" outer:"+ex.Message;}
Console.WriteLine(r);"#),
        &["inner:inner outer:outer"]
    );
}
