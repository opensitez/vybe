/// Pointer arithmetic and pointer walking — distinct from address/deref in test_pointers.rs.
use super::helpers::run_pascal;

#[test]
fn pointer_inc_walks_integer_array_forward() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..3] of Integer; p: ^Integer; begin a[0]:=1; a[1]:=2; a[2]:=3; a[3]:=4; p:=@a[0]; WriteLn(p^); Inc(p); WriteLn(p^); Inc(p); WriteLn(p^); end."#
        ),
        &["1", "2", "3"]
    );
}

#[test]
fn pointer_dec_walks_integer_array_backward() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..2] of Integer; p: ^Integer; begin a[2]:=30; p:=@a[2]; WriteLn(p^); Dec(p); WriteLn(p^); end."#
        ),
        &["30", "20"]
    );
}

#[test]
fn pointer_inc_twice_skips_one_element() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..3] of Integer; p: ^Integer; begin a[0]:=10; a[2]:=30; p:=@a[0]; Inc(p); Inc(p); WriteLn(p^); end."#
        ),
        &["30"]
    );
}

#[test]
fn pointer_starts_at_middle_element() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..5] of Integer; p: ^Integer; begin a[3]:=99; p:=@a[3]; WriteLn(p^); end."#
        ),
        &["99"]
    );
}

#[test]
fn pointer_loop_accumulates_via_inc() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..4] of Integer; p: ^Integer; i, s: Integer; begin for i:=0 to 4 do a[i]:=i+1; p:=@a[0]; s:=0; for i:=0 to 4 do begin s:=s+p^; Inc(p); end; WriteLn(s); end."#
        ),
        &["15"]
    );
}

#[test]
fn pointer_compare_same_element_address() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; p, q: ^Integer; begin x:=7; p:=@x; q:=@x; if p=q then WriteLn('same'); end."#
        ),
        &["same"]
    );
}

#[test]
fn pointer_compare_different_fields_not_equal() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record A,B:Integer; end; var r:TR; p,q:^Integer; begin r.A:=1; r.B:=2; p:=@r.A; q:=@r.B; if p<>q then WriteLn('diff'); end."#
        ),
        &["diff"]
    );
}

#[test]
fn pointer_write_through_inc_updates_next_cell() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..1] of Integer; p: ^Integer; begin a[0]:=0; a[1]:=0; p:=@a[0]; p^:=5; Inc(p); p^:=9; WriteLn(a[0]); WriteLn(a[1]); end."#
        ),
        &["5", "9"]
    );
}

#[test]
fn pointer_to_char_array_walks_string_buffer() {
    assert_eq!(
        run_pascal(
            r#"program T; var buf: array[0..2] of Char; p: ^Char; begin buf[0]:='a'; buf[1]:='b'; buf[2]:='c'; p:=@buf[0]; WriteLn(p^); Inc(p); WriteLn(p^); end."#
        ),
        &["a", "b"]
    );
}

#[test]
fn pointer_to_real_updates_through_deref() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Double; p: ^Double; begin x:=1.5; p:=@x; p^:=2.5; WriteLn(Round(p^*10)); end."#
        ),
        &["25"]
    );
}

#[test]
fn pointer_to_boolean_toggles_value() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: Boolean; p: ^Boolean; begin f:=false; p:=@f; p^:=true; if f then WriteLn('on'); end."#
        ),
        &["on"]
    );
}

#[test]
fn pointer_copy_between_pointer_vars() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; p,q: ^Integer; begin n:=42; p:=@n; q:=p; WriteLn(q^); end."#
        ),
        &["42"]
    );
}

#[test]
fn pointer_nil_before_assignment() {
    assert_eq!(
        run_pascal(
            r#"program T; var p: ^Integer; begin p:=nil; if not Assigned(p) then WriteLn('nil'); end."#
        ),
        &["nil"]
    );
}

#[test]
fn pointer_find_max_in_array_by_walking() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..3] of Integer; p: ^Integer; i, best: Integer; begin a[0]:=3; a[1]:=9; a[2]:=1; a[3]:=7; p:=@a[0]; best:=p^; for i:=1 to 3 do begin Inc(p); if p^>best then best:=p^; end; WriteLn(best); end."#
        ),
        &["9"]
    );
}

#[test]
fn pointer_swap_cells_using_temp() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,b: Integer; pa,pb: ^Integer; t: Integer; begin a:=1; b:=9; pa:=@a; pb:=@b; t:=pa^; pa^:=pb^; pb^:=t; WriteLn(a); WriteLn(b); end."#
        ),
        &["9", "1"]
    );
}

#[test]
fn pointer_offset_from_first_to_third_element() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..4] of Integer; first, third: ^Integer; begin a[0]:=10; a[2]:=30; first:=@a[0]; third:=@a[2]; WriteLn(first^); WriteLn(third^); end."#
        ),
        &["10", "30"]
    );
}

#[test]
fn pointer_passed_to_procedure_by_ref_walk() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Bump(var p: ^Integer); begin Inc(p); end; var a: array[0..1] of Integer; p: ^Integer; begin a[0]:=1; a[1]:=2; p:=@a[0]; Bump(p); WriteLn(p^); end."#
        ),
        &["2"]
    );
}

#[test]
fn pointer_returned_from_function_points_to_static() {
    assert_eq!(
        run_pascal(
            r#"program T; var g: Integer; function AddrGlobal: ^Integer; begin Result:=@g; end; var p: ^Integer; begin g:=55; p:=AddrGlobal; WriteLn(p^); end."#
        ),
        &["55"]
    );
}

#[test]
fn pointer_to_nested_record_field() {
    assert_eq!(
        run_pascal(
            r#"program T; type TInner=record V:Integer; end; type TOuter=record Inner:TInner; end; var o:TOuter; p:^Integer; begin o.Inner.V:=12; p:=@o.Inner.V; WriteLn(p^); end."#
        ),
        &["12"]
    );
}

#[test]
fn pointer_walk_counts_zero_cells() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..4] of Integer; p: ^Integer; i, c: Integer; begin for i:=0 to 4 do a[i]:=0; p:=@a[0]; c:=0; for i:=0 to 4 do begin if p^=0 then c:=c+1; Inc(p); end; WriteLn(c); end."#
        ),
        &["5"]
    );
}

#[test]
fn pointer_dec_after_inc_returns_original() {
    assert_eq!(
        run_pascal(
            r#"program T; var x: Integer; p: ^Integer; begin x:=8; p:=@x; Inc(p); Dec(p); WriteLn(p^); end."#
        ),
        &["8"]
    );
}

#[test]
fn pointer_subtract_same_array_yields_index_distance() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..5] of Integer; p, q: ^Integer; d: Integer; begin p:=@a[1]; q:=@a[4]; d:=q-p; WriteLn(d); end."#
        ),
        &["3"]
    );
}

#[test]
fn pointer_subtract_adjacent_cells_is_one() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..2] of Integer; p, q: ^Integer; begin p:=@a[0]; q:=@a[1]; WriteLn(q-p); end."#
        ),
        &["1"]
    );
}

#[test]
fn pointer_subtract_same_address_is_zero() {
    assert_eq!(
        run_pascal(
            r#"program T; var n: Integer; p, q: ^Integer; begin n:=0; p:=@n; q:=@n; WriteLn(q-p); end."#
        ),
        &["0"]
    );
}

#[test]
fn pointer_with_one_based_array_bounds() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..3] of Integer; p: ^Integer; begin a[2]:=44; p:=@a[2]; WriteLn(p^); end."#
        ),
        &["44"]
    );
}

#[test]
fn pointer_inc_in_while_until_end_index() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..2] of Integer; p: ^Integer; n: Integer; begin a[0]:=1; a[1]:=2; a[2]:=3; p:=@a[0]; n:=0; while n<3 do begin n:=n+p^; Inc(p); end; WriteLn(n); end."#
        ),
        &["6"]
    );
}

#[test]
fn pointer_parameter_reads_through_deref() {
    assert_eq!(
        run_pascal(
            r#"program T; function ReadPtr(p: ^Integer): Integer; begin Result:=p^; end; var x: Integer; begin x:=18; WriteLn(ReadPtr(@x)); end."#
        ),
        &["18"]
    );
}

#[test]
fn pointer_to_static_array_first_and_last() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..4] of Integer; first, last: ^Integer; begin a[0]:=2; a[4]:=8; first:=@a[0]; last:=@a[4]; WriteLn(last-first); end."#
        ),
        &["4"]
    );
}

#[test]
fn pointer_boolean_array_walk_finds_true() {
    assert_eq!(
        run_pascal(
            r#"program T; var f: array[0..2] of Boolean; p: ^Boolean; begin f[0]:=false; f[1]:=true; f[2]:=false; p:=@f[0]; Inc(p); if p^ then WriteLn('hit'); end."#
        ),
        &["hit"]
    );
}

#[test]
fn pointer_record_array_element_address() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=record V:Integer; end; var items: array[0..1] of TR; p: ^Integer; begin items[1].V:=77; p:=@items[1].V; WriteLn(p^); end."#
        ),
        &["77"]
    );
}

#[test]
fn pointer_assign_from_address_of_expression_index() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[1..4] of Integer; p: ^Integer; i: Integer; begin for i:=1 to 4 do a[i]:=i*10; p:=@a[3]; WriteLn(p^); end."#
        ),
        &["30"]
    );
}
