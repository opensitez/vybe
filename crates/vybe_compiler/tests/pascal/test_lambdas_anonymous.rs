/// Lambda-style anonymous methods and inline procedural expressions.
use super::helpers::run_pascal;

#[test]
fn anonymous_procedure_no_capture() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('anon'); end; p; end."#
        ),
        &["anon"]
    );
}

#[test]
fn anonymous_function_returns_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function:Integer; begin f:=function:Integer begin Result:=99; end; WriteLn(f()); end."#
        ),
        &["99"]
    );
}

#[test]
fn anonymous_method_with_one_param() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(n:Integer):Integer; begin f:=function(n:Integer):Integer begin Result:=n*n; end; WriteLn(f(6)); end."#
        ),
        &["36"]
    );
}

#[test]
fn anonymous_procedure_assigned_to_variable() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:procedure; begin a:=procedure begin WriteLn('a'); end; b:=a; b; end."#
        ),
        &["a"]
    );
}

#[test]
fn anonymous_function_in_array_map() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Map(const a:array of Integer; fn:function(n:Integer):Integer); var i:Integer; begin for i:=Low(a) to High(a) do WriteLn(fn(a[i])); end; begin Map([1,2,3], function(n:Integer):Integer begin Result:=n+10; end); end."#
        ),
        &["11", "12", "13"]
    );
}

#[test]
fn anonymous_nested_call() {
    assert_eq!(
        run_pascal(
            r#"program T; function Apply(fn:function(x:Integer):Integer; v:Integer):Integer; begin Result:=fn(v); end; begin WriteLn(Apply(function(x:Integer):Integer begin Result:=x*2; end, 4)); end."#
        ),
        &["8"]
    );
}

#[test]
fn anonymous_method_as_callback_filter() {
    assert_eq!(
        run_pascal(
            r#"program T; function CountIf(const a:array of Integer; pred:function(n:Integer):Boolean):Integer; var i:Integer; begin Result:=0; for i:=Low(a) to High(a) do if pred(a[i]) then Inc(Result); end; begin WriteLn(CountIf([1,2,3,4], function(n:Integer):Boolean begin Result:=n mod 2=0; end)); end."#
        ),
        &["2"]
    );
}

#[test]
fn anonymous_procedure_in_try_finally() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('run'); end; try p; finally WriteLn('done'); end; end."#
        ),
        &["run", "done"]
    );
}

#[test]
fn anonymous_function_string_concat() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(s:string):string; begin f:=function(s:string):string begin Result:=s+'!'; end; WriteLn(f('go')); end."#
        ),
        &["go!"]
    );
}

#[test]
fn anonymous_method_stored_in_record() {
    assert_eq!(
        run_pascal(
            r#"program T; type TWrap=record Fn:function:Integer; end; var w:TWrap; begin w.Fn:=function:Integer begin Result:=5; end; WriteLn(w.Fn()); end."#
        ),
        &["5"]
    );
}

#[test]
fn anonymous_procedure_twice_invocation() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin p:=procedure begin WriteLn('x'); end; p; p; end."#
        ),
        &["x", "x"]
    );
}

#[test]
fn anonymous_function_boolean_result() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(n:Integer):Boolean; begin f:=function(n:Integer):Boolean begin Result:=n>0; end; WriteLn(f(1)); WriteLn(f(-1)); end."#
        ),
        &["true", "false"]
    );
}

#[test]
fn anonymous_method_passed_to_higher_order() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Twice(fn:procedure); begin fn; fn; end; begin Twice(procedure begin WriteLn('t'); end); end."#
        ),
        &["t", "t"]
    );
}

#[test]
fn anonymous_function_closure_over_local() {
    assert_eq!(
        run_pascal(
            r#"program T; function MakeAdder(k:Integer):function(n:Integer):Integer; begin Result:=function(n:Integer):Integer begin Result:=n+k; end; end; var f:function(n:Integer):Integer; begin f:=MakeAdder(3); WriteLn(f(4)); end."#
        ),
        &["7"]
    );
}

#[test]
fn anonymous_method_in_class_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBox=class public Fn:procedure; end; var b:TBox; begin b:=TBox.Create; b.Fn:=procedure begin WriteLn('inbox'); end; b.Fn; b.Free; end."#
        ),
        &["inbox"]
    );
}

#[test]
fn anonymous_function_real_arithmetic() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(a,b:Double):Double; begin f:=function(a,b:Double):Double begin Result:=a+b; end; WriteLn(Round(f(1.2,2.3))); end."#
        ),
        &["4"]
    );
}

#[test]
fn anonymous_procedure_conditional_invoke() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; ok:Boolean; begin ok:=true; p:=procedure begin WriteLn('yes'); end; if ok then p; end."#
        ),
        &["yes"]
    );
}

#[test]
fn anonymous_function_compose_two() {
    assert_eq!(
        run_pascal(
            r#"program T; function Compose(f,g:function(n:Integer):Integer):function(n:Integer):Integer; begin Result:=function(n:Integer):Integer begin Result:=f(g(n)); end; end; var h:function(n:Integer):Integer; begin h:=Compose(function(n:Integer):Integer begin Result:=n+1; end, function(n:Integer):Integer begin Result:=n*2; end); WriteLn(h(3)); end."#
        ),
        &["7"]
    );
}

#[test]
fn anonymous_method_nil_check_before_call() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:procedure; begin if Assigned(p) then p else WriteLn('nil'); end."#
        ),
        &["nil"]
    );
}

#[test]
fn anonymous_function_char_transform() {
    assert_eq!(
        run_pascal(
            r#"program T; var f:function(c:Char):Char; begin f:=function(c:Char):Char begin Result:=UpCase(c); end; WriteLn(f('z')); end."#
        ),
        &["Z"]
    );
}
