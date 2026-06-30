/// Variable scope: local, nested, shadowing, block scope, unit-level.
use super::helpers::run_pascal;

#[test]
fn local_var_shadows_outer_same_name() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; procedure P; var x:Integer; begin x:=2; WriteLn(x); end; begin x:=1; P; WriteLn(x); end."#
        ),
        &["2", "1"]
    );
}

#[test]
fn nested_procedure_sees_outer_var() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; procedure Outer; procedure Inner; begin n:=n+1; end; begin Inner; end; begin n:=0; Outer; WriteLn(n); end."#
        ),
        &["1"]
    );
}

#[test]
fn block_local_var_lifetime() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:Integer; begin s:=0; begin var t:Integer; t:=5; s:=t; end; WriteLn(s); end."#
        ),
        &["5"]
    );
}

#[test]
fn for_loop_var_scope_inside_loop() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; begin i:=99; for i:=1 to 2 do WriteLn(i); WriteLn(i); end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn procedure_param_shadows_global() {
    assert_eq!(
        run_pascal(
            r#"program T; var v:Integer; procedure P(v:Integer); begin WriteLn(v); end; begin v:=1; P(9); end."#
        ),
        &["9"]
    );
}

#[test]
fn function_result_before_assign() {
    assert_eq!(
        run_pascal(
            r#"program T; function F:Integer; begin Result:=0; Result:=Result+4; end; begin WriteLn(F); end."#
        ),
        &["4"]
    );
}

#[test]
fn nested_function_accesses_outer_local() {
    assert_eq!(
        run_pascal(
            r#"program T; function Outer:Integer; var base:Integer; function Inner:Integer; begin Result:=base+1; end; begin base:=10; Result:=Inner; end; begin WriteLn(Outer); end."#
        ),
        &["11"]
    );
}

#[test]
fn const_in_inner_scope() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; const K=7; begin WriteLn(K); end; begin P; end."#
        ),
        &["7"]
    );
}

#[test]
fn typed_var_in_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; var s:string; begin s:='ok'; WriteLn(s); end; begin P; end."#
        ),
        &["ok"]
    );
}

#[test]
fn record_var_field_scope() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record A,B:Integer; end; var r:TR; begin r.A:=1; r.B:=2; WriteLn(r.A+r.B); end."#
        ),
        &["3"]
    );
}

#[test]
fn class_field_vs_local_name() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBox=class public V:Integer; procedure Show; end; procedure TBox.Show; var V:Integer; begin V:=5; WriteLn(Self.V); end; var b:TBox; begin b:=TBox.Create; b.V:=9; b.Show; b.Free; end."#
        ),
        &["9"]
    );
}

#[test]
fn unit_level_procedure_forward() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure B; forward; procedure A; begin B; end; procedure B; begin WriteLn('b'); end; begin A; end."#
        ),
        &["b"]
    );
}

#[test]
fn mutually_recursive_procedures() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Even(n:Integer); forward; procedure Odd(n:Integer); begin if n=0 then WriteLn('odd0') else Even(n-1); end; procedure Even(n:Integer); begin if n=0 then WriteLn('even0') else Odd(n-1); end; begin Even(2); end."#
        ),
        &["even0"]
    );
}

#[test]
fn var_section_multiple_declarations() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b,c:Integer; begin a:=1; b:=2; c:=3; WriteLn(a+b+c); end."#
        ),
        &["6"]
    );
}

#[test]
fn inline_var_declaration_in_begin() {
    assert_eq!(
        run_pascal(
            r#"program T; begin var x:Integer; x:=8; WriteLn(x); end."#
        ),
        &["8"]
    );
}

#[test]
fn outer_var_mutated_by_nested_proc() {
    assert_eq!(
        run_pascal(
            r#"program T; var total:Integer; procedure Acc(n:Integer); begin total:=total+n; end; begin total:=0; Acc(3); Acc(4); WriteLn(total); end."#
        ),
        &["7"]
    );
}

#[test]
fn case_block_local_scope() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=2; case n of 1: WriteLn('one'); 2: begin var t:Integer; t:=20; WriteLn(t); end; end; end."#
        ),
        &["20"]
    );
}

#[test]
fn repeat_loop_var_visible_after() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; begin i:=0; repeat Inc(i); until i=2; WriteLn(i); end."#
        ),
        &["2"]
    );
}

#[test]
fn while_loop_declared_before() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; begin n:=3; while n>0 do begin WriteLn(n); Dec(n); end; end."#
        ),
        &["3", "2", "1"]
    );
}

#[test]
fn nested_begin_end_isolation() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; begin x:=1; begin x:=2; end; WriteLn(x); end."#
        ),
        &["1"]
    );
}
