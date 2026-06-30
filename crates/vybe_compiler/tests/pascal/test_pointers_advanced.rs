/// Advanced pointers: New/Dispose, typed pointers, address-of variants.
use super::helpers::run_pascal;

#[test]
fn new_dispose_boolean_on_heap() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Boolean; begin New(p); p^:=true; WriteLn(p^); Dispose(p); end."#
        ),
        &["true"]
    );
}

#[test]
fn new_dispose_char_cell() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Char; begin New(p); p^:='X'; WriteLn(p^); Dispose(p); end."#
        ),
        &["X"]
    );
}

#[test]
fn new_dispose_double_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Double; begin New(p); p^:=2.5; WriteLn(p^>2.0); Dispose(p); end."#
        ),
        &["true"]
    );
}

#[test]
fn address_of_local_then_deref_read() {
    assert_eq!(
        run_pascal(
            r#"program T; var n:Integer; p:^Integer; begin n:=17; p:=@n; WriteLn(p^); end."#
        ),
        &["17"]
    );
}

#[test]
fn address_of_record_whole_struct() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record A,B:Integer; end; var r:TR; p:^TR; begin r.A:=2; r.B:=3; p:=@r; WriteLn(p^.A+p^.B); end."#
        ),
        &["5"]
    );
}

#[test]
fn address_of_array_element_middle() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[0..2] of Integer; p:^Integer; begin a[1]:=42; p:=@a[1]; WriteLn(p^); end."#
        ),
        &["42"]
    );
}

#[test]
fn typed_pointer_to_record_field_via_at() {
    assert_eq!(
        run_pascal(
            r#"program T; type TPt=record X,Y:Integer; end; var pt:TPt; px:^Integer; begin pt.Y:=6; px:=@pt.Y; px^:=8; WriteLn(pt.Y); end."#
        ),
        &["8"]
    );
}

#[test]
fn pointer_assign_between_same_typed_vars() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; p,q:^Integer; begin x:=3; p:=@x; q:=p; WriteLn(q^); end."#
        ),
        &["3"]
    );
}

#[test]
fn new_dispose_then_new_reuses_variable() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); p^:=1; Dispose(p); New(p); p^:=2; WriteLn(p^); Dispose(p); end."#
        ),
        &["2"]
    );
}

#[test]
fn pointer_to_string_concat_via_deref() {
    assert_eq!(
        run_pascal(
            r#"program T; var s:String; p:^String; begin s:='hi'; p:=@s; p^:=p^+'!'; WriteLn(s); end."#
        ),
        &["hi!"]
    );
}

#[test]
fn getmem_integer_via_pinteger_cast() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:Pointer; begin GetMem(p,SizeOf(Integer)); PInteger(p)^:=123; WriteLn(PInteger(p)^); FreeMem(p); end."#
        ),
        &["123"]
    );
}

#[test]
fn pointer_increment_across_two_cells() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[0..1] of Integer; p:^Integer; begin a[0]:=10; a[1]:=20; p:=@a[0]; Inc(p); WriteLn(p^); end."#
        ),
        &["20"]
    );
}

#[test]
fn pointer_decrement_across_two_cells() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[0..1] of Integer; p:^Integer; begin a[0]:=7; a[1]:=9; p:=@a[1]; Dec(p); WriteLn(p^); end."#
        ),
        &["7"]
    );
}

#[test]
fn new_dispose_record_on_heap() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; end; var p:^TR; begin New(p); p^.V:=99; WriteLn(p^.V); Dispose(p); end."#
        ),
        &["99"]
    );
}

#[test]
fn pointer_param_updates_caller_int() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure SetTen(p:^Integer); begin p^:=10; end; var x:Integer; begin x:=0; SetTen(@x); WriteLn(x); end."#
        ),
        &["10"]
    );
}

#[test]
fn function_returns_pointer_to_static() {
    assert_eq!(
        run_pascal(
            r#"program T; var G:Integer; function AddrG:^Integer; begin Result:=@G; end; var p:^Integer; begin G:=5; p:=AddrG; p^:=6; WriteLn(G); end."#
        ),
        &["6"]
    );
}

#[test]
fn double_pointer_assign_inner() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; p:^Integer; pp:^^Integer; begin x:=1; p:=@x; pp:=@p; pp^^:=4; WriteLn(x); end."#
        ),
        &["4"]
    );
}

#[test]
fn pointer_not_equal_after_retarget() {
    assert_eq!(
        run_pascal(
            r#"program T; var x,y:Integer; p:^Integer; begin x:=1; y:=2; p:=@x; p:=@y; WriteLn(p^); end."#
        ),
        &["2"]
    );
}

#[test]
fn new_dispose_in_loop_three_times() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; i,sum:Integer; begin sum:=0; for i:=1 to 3 do begin New(p); p^:=i; sum:=sum+p^; Dispose(p); end; WriteLn(sum); end."#
        ),
        &["6"]
    );
}

#[test]
fn pointer_to_byte_updates_array_slot() {
    assert_eq!(
        run_pascal(
            r#"program T; var b:Byte; p:^Byte; begin b:=0; p:=@b; p^:=255; WriteLn(b); end."#
        ),
        &["255"]
    );
}

#[test]
fn typed_pointer_subtract_returns_difference() {
    assert_eq!(
        run_pascal(
            r#"program T; var a:array[0..2] of Integer; p,q:^Integer; begin a[0]:=1; a[2]:=3; p:=@a[0]; q:=@a[2]; WriteLn((q-p)>=0); end."#
        ),
        &["true"]
    );
}

#[test]
fn pointer_boolean_negate_via_deref() {
    assert_eq!(
        run_pascal(
            r#"program T; var b:Boolean; p:^Boolean; begin b:=true; p:=@b; p^:=not p^; WriteLn(b); end."#
        ),
        &["false"]
    );
}

#[test]
fn new_dispose_string_pointer() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^String; begin New(p); p^:='heap'; WriteLn(p^); Dispose(p); end."#
        ),
        &["heap"]
    );
}

#[test]
fn address_of_for_loop_index_var() {
    assert_eq!(
        run_pascal(
            r#"program T; var i:Integer; p:^Integer; begin for i:=1 to 1 do begin p:=@i; WriteLn(p^); end; end."#
        ),
        &["1"]
    );
}

#[test]
fn pointer_passed_through_nested_call() {
    assert_eq!(
        run_pascal(
            r#"program T; function Id(p:^Integer):^Integer; begin Result:=p; end; var x:Integer; p:^Integer; begin x:=8; p:=Id(@x); WriteLn(p^); end."#
        ),
        &["8"]
    );
}

#[test]
fn dispose_after_mutate_multiple_fields() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record A,B:Integer; end; var p:^TR; begin New(p); p^.A:=1; p^.B:=2; WriteLn(p^.A+p^.B); Dispose(p); WriteLn('done'); end."#
        ),
        &["3", "done"]
    );
}

#[test]
fn pointer_compare_nil_after_dispose_var() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^Integer; begin New(p); Dispose(p); WriteLn('ok'); end."#
        ),
        &["ok"]
    );
}

#[test]
fn pchar_style_char_pointer_deref() {
    assert_eq!(
        run_pascal(
            r#"program T; var c:Char; p:^Char; begin c:='M'; p:=@c; WriteLn(p^); end."#
        ),
        &["M"]
    );
}

#[test]
fn pointer_to_word_arithmetic_field() {
    assert_eq!(
        run_pascal(
            r#"program T; var w:Word; p:^Word; begin w:=100; p:=@w; p^:=p^+50; WriteLn(w); end."#
        ),
        &["150"]
    );
}

#[test]
fn new_dispose_smallint_negative() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^SmallInt; begin New(p); p^:=-12; WriteLn(p^); Dispose(p); end."#
        ),
        &["-12"]
    );
}

#[test]
fn pointer_equality_after_reassign_same_addr() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; p,q:^Integer; begin x:=0; p:=@x; q:=@x; WriteLn(p=q); end."#
        ),
        &["true"]
    );
}

#[test]
fn pointer_inequality_different_locals() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b:Integer; p,q:^Integer; begin a:=1; b:=1; p:=@a; q:=@b; WriteLn(p<>q); end."#
        ),
        &["true"]
    );
}

#[test]
fn getmem_zero_initialize_then_set() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:Pointer; begin GetMem(p,SizeOf(Integer)); PInteger(p)^:=0; Inc(PInteger(p)^); WriteLn(PInteger(p)^); FreeMem(p); end."#
        ),
        &["1"]
    );
}

#[test]
fn pointer_to_nested_record_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=record V:Integer; end; TOuter=record N:TInner; end; var o:TOuter; p:^Integer; begin o.N.V:=4; p:=@o.N.V; p^:=5; WriteLn(o.N.V); end."#
        ),
        &["5"]
    );
}

#[test]
fn new_dispose_two_separate_pointers() {
    assert_eq!(
        run_pascal(
            r#"program T; var p,q:^Integer; begin New(p); New(q); p^:=1; q^:=2; WriteLn(p^+q^); Dispose(p); Dispose(q); end."#
        ),
        &["3"]
    );
}

#[test]
fn address_of_function_result_via_temp() {
    assert_eq!(
        run_pascal(
            r#"program T; function Make:Integer; begin Result:=11; end; var n:Integer; p:^Integer; begin n:=Make; p:=@n; WriteLn(p^); end."#
        ),
        &["11"]
    );
}

#[test]
fn pointer_read_modify_write_expr() {
    assert_eq!(
        run_pascal(
            r#"program T; var x:Integer; p:^Integer; begin x:=2; p:=@x; p^:=p^*p^; WriteLn(x); end."#
        ),
        &["4"]
    );
}

#[test]
fn typed_pointer_in_record_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TWrap=record P:^Integer; end; var x:Integer; w:TWrap; begin x:=9; w.P:=@x; WriteLn(w.P^); end."#
        ),
        &["9"]
    );
}

#[test]
fn new_dispose_longint_large_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var p:^LongInt; begin New(p); p^:=100000; WriteLn(p^); Dispose(p); end."#
        ),
        &["100000"]
    );
}

#[test]
fn pointer_param_const_does_not_change_pointer() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Show(p:^Integer); begin WriteLn(p^); end; var x:Integer; begin x:=44; Show(@x); end."#
        ),
        &["44"]
    );
}
