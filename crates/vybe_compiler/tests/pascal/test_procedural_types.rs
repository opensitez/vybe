/// Procedural types and anonymous methods — delegates/lambdas as values.
use super::helpers::run_pascal;

#[test]
fn procedural_type_function_variable_stored() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: function(x: Integer): Integer; begin f:=function(x: Integer): Integer begin Result:=x+1; end; WriteLn(f(4)); end."#
        ),
        &["5"]
    );
}

#[test]
fn procedural_type_procedure_variable_invoked() {
    assert_eq!(
        run_pascal(
            r#"program T; var p: procedure; begin p:=procedure begin WriteLn('run'); end; p; end."#
        ),
        &["run"]
    );
}

#[test]
fn procedural_type_passed_as_parameter() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Twice(f: function(x: Integer): Integer; n: Integer); begin WriteLn(f(n)); WriteLn(f(n)); end; begin Twice(function(x: Integer): Integer begin Result:=x*2; end, 3); end."#
        ),
        &["6", "6"]
    );
}

#[test]
fn procedural_type_returns_from_function() {
    assert_eq!(
        run_pascal(
            r#"program T; function MakeAdder(k: Integer): function(x: Integer): Integer; begin Result:=function(x: Integer): Integer begin Result:=x+k; end; end; var f: function(x: Integer): Integer; begin f:=MakeAdder(10); WriteLn(f(5)); end."#
        ),
        &["15"]
    );
}

#[test]
fn procedural_type_procedure_with_string_arg() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Echo(p: procedure(s: String)); begin p('hi'); end; begin Echo(procedure(s: String) begin WriteLn(s); end); end."#
        ),
        &["hi"]
    );
}

#[test]
fn procedural_type_nested_call_chain() {
    assert_eq!(
        run_pascal(
            r#"program T; function Inc1: function(x: Integer): Integer; begin Result:=function(x: Integer): Integer begin Result:=x+1; end; end; var f: function(x: Integer): Integer; begin f:=Inc1; WriteLn(f(f(0))); end."#
        ),
        &["2"]
    );
}

#[test]
fn procedural_type_capture_outer_local_in_anonymous() {
    assert_eq!(
        run_pascal(
            r#"program T; var base: Integer; f: function(x: Integer): Integer; begin base:=100; f:=function(x: Integer): Integer begin Result:=base+x; end; WriteLn(f(5)); end."#
        ),
        &["105"]
    );
}

#[test]
fn procedural_type_reassign_changes_behavior() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: function(x: Integer): Integer; begin f:=function(x: Integer): Integer begin Result:=x; end; WriteLn(f(1)); f:=function(x: Integer): Integer begin Result:=x*10; end; WriteLn(f(1)); end."#
        ),
        &["1", "10"]
    );
}

#[test]
fn procedural_type_boolean_predicate() {
    assert_eq!(
        run_pascal(
            r#"program T; var ok: function(n: Integer): Boolean; begin ok:=function(n: Integer): Boolean begin Result:=n mod 2=0; end; if ok(8) then WriteLn('even'); end."#
        ),
        &["even"]
    );
}

#[test]
fn procedural_type_filter_with_callback() {
    assert_eq!(
        run_pascal(
            r#"program T; function CountIf(a: array of Integer; pred: function(x: Integer): Boolean): Integer; var i: Integer; begin Result:=0; for i:=0 to High(a) do if pred(a[i]) then Result:=Result+1; end; begin WriteLn(CountIf([1,2,3,4], function(x: Integer): Boolean begin Result:=x>2; end)); end."#
        ),
        &["2"]
    );
}

#[test]
fn procedural_type_map_with_callback() {
    assert_eq!(
        run_pascal(
            r#"program T; function Apply(a: array of Integer; fn: function(x: Integer): Integer): Integer; begin Result:=fn(a[0])+fn(a[1]); end; begin WriteLn(Apply([3,4], function(x: Integer): Integer begin Result:=x*x; end)); end."#
        ),
        &["25"]
    );
}

#[test]
fn procedural_type_procedure_closure_counter() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; tick: procedure; begin n:=0; tick:=procedure begin n:=n+1; end; tick; tick; WriteLn(n); end."#
        ),
        &["2"]
    );
}

#[test]
fn procedural_type_is_nil_before_assign() {
    assert_eq!(
        run_pascal(
            r#"program T; var p: procedure; begin if not Assigned(p) then WriteLn('unset'); p:=procedure begin WriteLn('set'); end; p; end."#
        ),
        &["unset", "set"]
    );
}

#[test]
fn procedural_type_method_style_on_record_helper_like() {
    assert_eq!(
        run_pascal(
            r#"program T; type TFn=function(x: Integer): Integer; function DoubleFn: TFn; begin Result:=function(x: Integer): Integer begin Result:=x*2; end; end; var f: TFn; begin f:=DoubleFn; WriteLn(f(6)); end."#
        ),
        &["12"]
    );
}

#[test]
fn procedural_type_compose_two_functions() {
    assert_eq!(
        run_pascal(
            r#"program T; function Compose(f,g: function(x: Integer): Integer; v: Integer): Integer; begin Result:=f(g(v)); end; begin WriteLn(Compose(function(x: Integer): Integer begin Result:=x+1; end, function(x: Integer): Integer begin Result:=x*2; end, 3)); end."#
        ),
        &["7"]
    );
}

#[test]
fn procedural_type_string_transform_callback() {
    assert_eq!(
        run_pascal(
            r#"program T; function Transform(s: String; fn: function(c: Char): Char): String; var i: Integer; begin Result:=''; for i:=1 to Length(s) do Result:=Result+fn(s[i]); end; begin WriteLn(Transform('ab', function(c: Char): Char begin Result:=UpCase(c); end)); end."#
        ),
        &["AB"]
    );
}

#[test]
fn procedural_type_reduce_via_procedure_var() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Fold(var acc: Integer; add: procedure(var a: Integer; v: Integer); v: Integer); begin add(acc,v); end; var total: Integer; begin total:=0; Fold(total, procedure(var a: Integer; v: Integer) begin a:=a+v; end, 5); Fold(total, procedure(var a: Integer; v: Integer) begin a:=a+v; end, 7); WriteLn(total); end."#
        ),
        &["12"]
    );
}

#[test]
fn procedural_type_compare_callback_sort_key() {
    assert_eq!(
        run_pascal(
            r#"program T; function PickMax(a,b: Integer; better: function(x,y: Integer): Boolean): Integer; begin if better(a,b) then Result:=a else Result:=b; end; begin WriteLn(PickMax(3,9, function(x,y: Integer): Boolean begin Result:=x>y; end)); end."#
        ),
        &["9"]
    );
}

#[test]
fn procedural_type_void_observer_pattern() {
    assert_eq!(
        run_pascal(
            r#"program T; var obs: procedure; procedure Notify; begin if Assigned(obs) then obs; end; begin obs:=procedure begin WriteLn('ping'); end; Notify; end."#
        ),
        &["ping"]
    );
}

#[test]
fn procedural_type_reference_parameter_in_callback() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Mutate(var n: Integer; step: procedure(var x: Integer)); begin step(n); end; var v: Integer; begin v:=5; Mutate(v, procedure(var x: Integer) begin x:=x*3; end); WriteLn(v); end."#
        ),
        &["15"]
    );
}

#[test]
fn procedural_type_nested_anonymous_in_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; var i: Integer; f: function(): Integer; s: Integer; begin s:=0; for i:=1 to 3 do begin f:=function(): Integer begin Result:=i; end; s:=s+f; end; WriteLn(s); end."#
        ),
        &["6"]
    );
}

#[test]
fn procedural_type_event_handler_two_listeners() {
    assert_eq!(
        run_pascal(
            r#"program T; var h1,h2: procedure; procedure Fire; begin h1; h2; end; begin h1:=procedure begin WriteLn('one'); end; h2:=procedure begin WriteLn('two'); end; Fire; end."#
        ),
        &["one", "two"]
    );
}

#[test]
fn procedural_type_function_returns_string() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: function(n: Integer): String; begin f:=function(n: Integer): String begin Result:=IntToStr(n)+'!'; end; WriteLn(f(7)); end."#
        ),
        &["7!"]
    );
}

#[test]
fn procedural_type_dispatch_table_by_index() {
    assert_eq!(
        run_pascal(
            r#"program T; var table: array[0..1] of function(x: Integer): Integer; begin table[0]:=function(x: Integer): Integer begin Result:=x; end; table[1]:=function(x: Integer): Integer begin Result:=x+10; end; WriteLn(table[1](5)); end."#
        ),
        &["15"]
    );
}

#[test]
fn procedural_type_lazy_factory() {
    assert_eq!(
        run_pascal(
            r#"program T; function Lazy: function(): Integer; begin Result:=function(): Integer begin Result:=99; end; end; var g: function(): Integer; begin g:=Lazy; WriteLn(g()); end."#
        ),
        &["99"]
    );
}
