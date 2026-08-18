/// Multidimensional static and dynamic arrays.
use super::helpers::run_pascal;

#[test]
fn matrix_2d_dimensions() {
    assert_eq!(
        run_pascal(
            r#"program T; var m: array[1..2,1..3] of Integer; begin WriteLn(Low(m,1)); WriteLn(High(m,2)); end."#
        ),
        &["1", "3"]
    );
}

#[test]
fn matrix_2d_store_and_read() {
    assert_eq!(
        run_pascal(
            r#"program T; var m: array[0..1,0..1] of Integer; begin m[0,0]:=1; m[1,1]:=4; WriteLn(m[0,0]+m[1,1]); end."#
        ),
        &["5"]
    );
}

#[test]
fn matrix_2d_row_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; var m: array[1..2,1..2] of Integer; r,c,s: Integer; begin m[1,1]:=1; m[1,2]:=2; m[2,1]:=3; m[2,2]:=4; s:=0; for r:=1 to 2 do for c:=1 to 2 do s:=s+m[r,c]; WriteLn(s); end."#
        ),
        &["10"]
    );
}

#[test]
fn matrix_2d_diagonal_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; var m: array[1..3,1..3] of Integer; i,s: Integer; begin for i:=1 to 3 do m[i,i]:=i; s:=0; for i:=1 to 3 do s:=s+m[i,i]; WriteLn(s); end."#
        ),
        &["6"]
    );
}

#[test]
fn matrix_2d_transpose_copy() {
    assert_eq!(
        run_pascal(
            r#"program T; var a,t: array[1..2,1..2] of Integer; r,c: Integer; begin a[1,2]:=9; for r:=1 to 2 do for c:=1 to 2 do t[c,r]:=a[r,c]; WriteLn(t[2,1]); end."#
        ),
        &["9"]
    );
}

#[test]
fn matrix_3d_index_and_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; var a: array[0..1,0..1,0..1] of Integer; x,y,z,s: Integer; begin for x:=0 to 1 do for y:=0 to 1 do for z:=0 to 1 do a[x,y,z]:=1; s:=0; for x:=0 to 1 do for y:=0 to 1 do for z:=0 to 1 do s:=s+a[x,y,z]; WriteLn(s); end."#
        ),
        &["8"]
    );
}

#[test]
fn matrix_nested_record_cell() {
    assert_eq!(
        run_pascal(
            r#"program T; type TC=record V:Integer; end; var g: array[1..2,1..2] of TC; begin g[1,1].V:=5; g[2,2].V:=6; WriteLn(g[1,1].V+g[2,2].V); end."#
        ),
        &["11"]
    );
}

#[test]
fn matrix_open_array_param_rows() {
    assert_eq!(
        run_pascal(
            r#"program T; function Cell(const m: array of array of Integer; r,c: Integer): Integer; begin Result:=m[r][c]; end; var m: array[0..1] of array[0..1] of Integer; begin m[0][0]:=3; m[1][1]:=4; WriteLn(Cell(m,1,1)); end."#
        ),
        &["4"]
    );
}

#[test]
fn matrix_char_grid_print_row() {
    assert_eq!(
        run_pascal(
            r#"program T; var g: array[1..2,1..2] of Char; c: Integer; begin g[1,1]:='a'; g[1,2]:='b'; for c:=1 to 2 do Write(g[1,c]); WriteLn; end."#
        ),
        &["ab"]
    );
}

#[test]
fn matrix_boolean_and_reduction() {
    assert_eq!(
        run_pascal(
            r#"program T; var m: array[1..2,1..2] of Boolean; r,c: Integer; ok: Boolean; begin for r:=1 to 2 do for c:=1 to 2 do m[r,c]:=true; ok:=m[1,1] and m[2,2]; WriteLn(ok); end."#
        ),
        &["TRUE"]
    );
}

#[test]
fn matrix_negative_bounds() {
    assert_eq!(
        run_pascal(
            r#"program T; var m: array[-1..0,-1..0] of Integer; begin m[-1,-1]:=2; m[0,0]:=3; WriteLn(m[-1,-1]+m[0,0]); end."#
        ),
        &["5"]
    );
}

#[test]
fn matrix_fill_sequence() {
    assert_eq!(
        run_pascal(
            r#"program T; var m: array[1..2,1..3] of Integer; r,c,v: Integer; begin v:=0; for r:=1 to 2 do for c:=1 to 3 do begin Inc(v); m[r,c]:=v; end; WriteLn(m[2,3]); end."#
        ),
        &["6"]
    );
}

#[test]
fn dynamic_matrix_setlength_rows() {
    assert_eq!(
        run_pascal(
            r#"program T; var m: array of array of Integer; begin SetLength(m,2); SetLength(m[0],2); SetLength(m[1],2); m[0][0]:=1; m[1][1]:=2; WriteLn(m[0][0]+m[1][1]); end."#
        ),
        &["3"]
    );
}

#[test]
fn dynamic_matrix_high_low_per_dim() {
    assert_eq!(
        run_pascal(
            r#"program T; var m: array of array of Integer; begin SetLength(m,2); SetLength(m[0],3); WriteLn(High(m)); WriteLn(High(m[0])); end."#
        ),
        &["1", "2"]
    );
}

#[test]
fn matrix_var_param_mutation() {
    assert_eq!(
        run_pascal(
            r#"program T; procedure Bump(var m: array[1..1,1..1] of Integer); begin m[1,1]:=m[1,1]+1; end; var a: array[1..1,1..1] of Integer; begin a[1,1]:=4; Bump(a); WriteLn(a[1,1]); end."#
        ),
        &["5"]
    );
}

#[test]
fn matrix_function_returns_row_sum() {
    assert_eq!(
        run_pascal(
            r#"program T; function RowSum(const m: array[1..2,1..2] of Integer; r: Integer): Integer; var c: Integer; begin Result:=0; for c:=1 to 2 do Result:=Result+m[r,c]; end; var m: array[1..2,1..2] of Integer; begin m[1,1]:=2; m[1,2]:=3; WriteLn(RowSum(m,1)); end."#
        ),
        &["5"]
    );
}

#[test]
fn matrix_swap_two_cells() {
    assert_eq!(
        run_pascal(
            r#"program T; var m: array[1..1,1..2] of Integer; t: Integer; begin m[1,1]:=1; m[1,2]:=9; t:=m[1,1]; m[1,1]:=m[1,2]; m[1,2]:=t; WriteLn(m[1,1]); end."#
        ),
        &["9"]
    );
}

#[test]
fn matrix_enum_as_first_dimension() {
    assert_eq!(
        run_pascal(
            r#"program T; type TR=(R1,R2); var m: array[TR,1..2] of Integer; d: TR; begin m[R1,1]:=7; m[R2,2]:=8; WriteLn(m[R1,1]+m[R2,2]); end."#
        ),
        &["15"]
    );
}

#[test]
fn matrix_class_field_2d() {
    assert_eq!(
        run_pascal(
            r#"program T; type TBoard=class public Cells: array[0..1,0..1] of Integer; end; var b: TBoard; begin b:=TBoard.Create; b.Cells[0,1]:=11; b.Cells[1,0]:=22; WriteLn(b.Cells[0,1]+b.Cells[1,0]); b.Free; end."#
        ),
        &["33"]
    );
}

#[test]
fn matrix_identity_check() {
    assert_eq!(
        run_pascal(
            r#"program T; var m: array[1..2,1..2] of Integer; r,c: Integer; ok: Boolean; begin for r:=1 to 2 do for c:=1 to 2 do if r=c then m[r,c]:=1 else m[r,c]:=0; ok:=(m[1,1]=1) and (m[2,2]=1) and (m[1,2]=0); WriteLn(ok); end."#
        ),
        &["TRUE"]
    );
}
