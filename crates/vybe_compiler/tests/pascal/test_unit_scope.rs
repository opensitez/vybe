/// Unit and nested scope: forward declarations, nested procedures, name hiding.
use super::helpers::run_pascal;

#[test]
fn forward_procedure_call_before_impl() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Alpha; forward; procedure Beta; begin WriteLn('beta'); end; procedure Alpha; begin Beta; end; begin Alpha; end."#
        ),
        &["beta"]
    );
}

#[test]
fn forward_function_call_before_impl() {
    assert_eq!(
        run_pascal(
            r#"program T; function DoubleIt(n: Integer): Integer; forward; function TripleIt(n: Integer): Integer; begin Result:=DoubleIt(n)+n; end; function DoubleIt(n: Integer): Integer; begin Result:=n*2; end; begin WriteLn(TripleIt(5)); end."#
        ),
        &["15"]
    );
}

#[test]
fn nested_procedure_reads_outer_local() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; var x: Integer; procedure Inner; begin WriteLn(x); end; begin x:=9; Inner; end; begin Outer; end."#
        ),
        &["9"]
    );
}

#[test]
fn nested_procedure_modifies_outer_local() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer; var x: Integer; procedure Inner; begin x:=x+1; end; begin x:=1; Inner; WriteLn(x); end; begin Outer; end."#
        ),
        &["2"]
    );
}

#[test]
fn inner_var_shadows_outer_var() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; procedure P; var x: Integer; begin x:=100; WriteLn(x); end; begin x:=1; P; WriteLn(x); end."#
        ),
        &["100", "1"]
    );
}

#[test]
fn triple_nested_procedure_chain() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure A; procedure B; procedure C; begin WriteLn('deep'); end; begin C; end; begin B; end; begin A; end."#
        ),
        &["deep"]
    );
}

#[test]
fn local_const_visible_in_nested_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; const K=7; procedure Q; begin WriteLn(K); end; begin Q; end; begin P; end."#
        ),
        &["7"]
    );
}

#[test]
fn local_type_used_in_nested_function() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; type TPair=record A,B:Integer; end; function Make: TPair; begin Result.A:=1; Result.B:=2; end; begin WriteLn(Make.B); end; begin P; end."#
        ),
        &["2"]
    );
}

#[test]
fn procedure_local_var_not_visible_outside() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; procedure P; var y: Integer; begin y:=5; x:=y; end; begin x:=0; P; WriteLn(x); end."#
        ),
        &["5"]
    );
}

#[test]
fn sibling_nested_procedures_share_outer() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; var s: Integer; procedure Add1; begin s:=s+1; end; procedure Add2; begin s:=s+2; end; begin s:=0; Add1; Add2; WriteLn(s); end; begin P; end."#
        ),
        &["3"]
    );
}

#[test]
fn global_var_visible_in_nested_two_levels() {
    assert_eq!(
        run_pascal(
            r#"program T; var g: Integer; procedure A; procedure B; begin g:=42; end; begin B; end; begin g:=0; A; WriteLn(g); end."#
        ),
        &["42"]
    );
}

#[test]
fn parameter_shadows_global_name_in_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; procedure Show(n: Integer); begin WriteLn(n); end; begin n:=1; Show(9); WriteLn(n); end."#
        ),
        &["9", "1"]
    );
}

#[test]
fn nested_procedure_passes_outer_to_inner_via_param() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Outer(v: Integer); procedure Inner(x: Integer); begin WriteLn(x); end; begin Inner(v); end; begin Outer(6); end."#
        ),
        &["6"]
    );
}

#[test]
fn function_result_visible_after_nested_call() {
    assert_eq!(
        run_pascal(
            r#"program T; function Calc: Integer; function Helper: Integer; begin Result:=4; end; begin Result:=Helper*3; end; begin WriteLn(Calc); end."#
        ),
        &["12"]
    );
}

#[test]
fn nested_procedure_early_exit_does_not_leak() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; var n: Integer; procedure Q; begin n:=3; Exit; n:=9; end; begin n:=0; Q; WriteLn(n); end; begin P; end."#
        ),
        &["3"]
    );
}

#[test]
fn outer_for_loop_var_visible_in_nested_proc() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; var i, s: Integer; procedure Acc; begin s:=s+i; end; begin s:=0; for i:=1 to 3 do Acc; WriteLn(s); end; begin P; end."#
        ),
        &["6"]
    );
}

#[test]
fn nested_procedure_with_local_array() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; var a: array[0..1] of Integer; procedure Fill; begin a[0]:=1; a[1]:=2; end; begin Fill; WriteLn(a[1]); end; begin P; end."#
        ),
        &["2"]
    );
}

#[test]
fn forward_mutual_procedure_a_calls_b() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure A; forward; procedure B; forward; procedure A; begin WriteLn('A'); B; end; procedure B; begin WriteLn('B'); end; begin A; end."#
        ),
        &["A", "B"]
    );
}

#[test]
fn nested_class_method_sees_unit_level_var() {
    assert_eq!(
        run_pascal(
            r#"program T; var total: Integer; type T=class procedure Bump; end; class procedure T.Bump; begin total:=total+1; end; begin total:=0; T.Bump; T.Bump; WriteLn(total); end."#
        ),
        &["2"]
    );
}

#[test]
fn local_record_scope_in_procedure() {
    assert_eq!(
        run_pascal(
            r#"program T; type TG=record V:Integer; end; var g:TG; procedure P; var r:TG; begin r.V:=8; g:=r; end; begin g.V:=0; P; WriteLn(g.V); end."#
        ),
        &["8"]
    );
}

#[test]
fn nested_procedure_string_builder() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; var s: string; procedure Append(c: Char); begin s:=s+c; end; begin s:=''; Append('a'); Append('b'); WriteLn(s); end; begin P; end."#
        ),
        &["ab"]
    );
}

#[test]
fn outer_boolean_flag_set_from_nested() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; var ok: Boolean; procedure SetOk; begin ok:=true; end; begin ok:=false; SetOk; if ok then WriteLn('yes'); end; begin P; end."#
        ),
        &["yes"]
    );
}

#[test]
fn nested_procedure_recursive_on_counter() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; var n: Integer; procedure Rec; begin if n<3 then begin n:=n+1; Rec; end; end; begin n:=0; Rec; WriteLn(n); end; begin P; end."#
        ),
        &["3"]
    );
}

#[test]
fn local_enum_in_procedure_scope() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; type TS=(On,Off); var s:TS; begin s:=On; if s=On then WriteLn('on'); end; begin P; end."#
        ),
        &["on"]
    );
}

#[test]
fn nested_function_returns_to_outer_assign() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; var v: Integer; function Inner: Integer; begin Result:=21; end; begin v:=Inner; WriteLn(v); end; begin P; end."#
        ),
        &["21"]
    );
}

#[test]
fn procedure_scope_preserves_global_const() {
    assert_eq!(
        run_pascal(
            r#"program T; const G=11; procedure P; const L=2; begin WriteLn(G+L); end; begin P; end."#
        ),
        &["13"]
    );
}

#[test]
fn nested_procedure_iterates_outer_array() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure P; var a: array[0..2] of Integer; procedure Sum(var s: Integer); var i: Integer; begin s:=0; for i:=0 to 2 do s:=s+a[i]; end; var t: Integer; begin a[0]:=1; a[1]:=2; a[2]:=3; Sum(t); WriteLn(t); end; begin P; end."#
        ),
        &["6"]
    );
}

#[test]
fn forward_function_mutual_odd_even() {
    assert_eq!(
        run_pascal(
            r#"program T; function IsEven(n: Integer): Boolean; forward; function IsOdd(n: Integer): Boolean; begin if n=0 then Result:=false else Result:=IsEven(n-1); end; function IsEven(n: Integer): Boolean; begin if n=0 then Result:=true else Result:=IsOdd(n-1); end; begin if IsEven(4) then WriteLn('even'); end."#
        ),
        &["even"]
    );
}

#[test]
fn nested_procedure_case_on_outer_enum() {
    assert_eq!(
        run_pascal(
            r#"program T; type TM=(A,B); procedure P; var m: TM; procedure Show; begin case m of A: WriteLn('A'); B: WriteLn('B'); end; end; begin m:=B; Show; end; begin P; end."#
        ),
        &["B"]
    );
}

#[test]
fn deeply_nested_var_init_once() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure L1; var n: Integer; procedure L2; procedure L3; begin n:=5; end; begin L3; WriteLn(n); end; begin n:=0; L2; end; begin L1; end."#
        ),
        &["5"]
    );
}
