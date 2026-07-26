use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 87: Multi-Dimensional Dynamic Arrays & Ragged Matrices
// ═══════════════════════════════════════════════════════════

#[test]
fn test_dynarray_2d_setlength_rectangular() {
    let out = run_pascal(
        r#"
program Test;
type TMatrix = array of array of Integer;
var m: TMatrix;
begin
  SetLength(m, 2, 3);
  WriteLn(Length(m));
  WriteLn(Length(m[0]));
end.
"#,
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn test_dynarray_2d_element_read_write() {
    let out = run_pascal(
        r#"
program Test;
type TMatrix = array of array of Integer;
var m: TMatrix;
begin
  SetLength(m, 2, 2);
  m[0, 0] := 1; m[0, 1] := 2;
  m[1, 0] := 3; m[1, 1] := 4;
  WriteLn(m[0, 0].ToString + ',' + m[0, 1].ToString);
  WriteLn(m[1, 0].ToString + ',' + m[1, 1].ToString);
end.
"#,
    );
    assert_eq!(out, vec!["1,2", "3,4"]);
}

#[test]
fn test_dynarray_2d_ragged_matrix() {
    let out = run_pascal(
        r#"
program Test;
type TRagged = array of array of Integer;
var r: TRagged;
begin
  SetLength(r, 3);
  SetLength(r[0], 1);
  SetLength(r[1], 2);
  SetLength(r[2], 3);

  WriteLn(Length(r[0]));
  WriteLn(Length(r[1]));
  WriteLn(Length(r[2]));
end.
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn test_dynarray_2d_bounds_low_high() {
    let out = run_pascal(
        r#"
program Test;
type TMatrix = array of array of Double;
var m: TMatrix;
begin
  SetLength(m, 3, 4);
  WriteLn(Low(m).ToString + '..' + High(m).ToString);
  WriteLn(Low(m[0]).ToString + '..' + High(m[0]).ToString);
end.
"#,
    );
    assert_eq!(out, vec!["0..2", "0..3"]);
}

#[test]
fn test_dynarray_3d_cube() {
    let out = run_pascal(
        r#"
program Test;
type TCube = array of array of array of Byte;
var c: TCube;
begin
  SetLength(c, 2, 3, 4);
  c[1, 2, 3] := 255;
  WriteLn(Length(c));
  WriteLn(Length(c[0]));
  WriteLn(Length(c[0, 0]));
  WriteLn(c[1, 2, 3]);
end.
"#,
    );
    assert_eq!(out, vec!["2", "3", "4", "255"]);
}

#[test]
fn test_dynarray_2d_reference_assignment() {
    let out = run_pascal(
        r#"
program Test;
type TMatrix = array of array of Integer;
var m1, m2: TMatrix;
begin
  SetLength(m1, 2, 2);
  m1[0, 0] := 42;
  m2 := m1; // Reference assignment
  WriteLn(m2[0, 0]);
  m1[0, 0] := 99;
  WriteLn(m2[0, 0]);
end.
"#,
    );
    assert_eq!(out, vec!["42", "99"]);
}

#[test]
fn test_dynarray_2d_copy_function() {
    let out = run_pascal(
        r#"
program Test;
type TMatrix = array of array of Integer;
var m1, m2: TMatrix;
begin
  SetLength(m1, 2, 2);
  m1[0, 0] := 10;
  m2 := Copy(m1);
  m1[0, 0] := 20;
  WriteLn(m2[0, 0]);
end.
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_dynarray_2d_nested_loop_iteration() {
    let out = run_pascal(
        r#"
program Test;
type TGrid = array of array of Integer;
var g: TGrid; r, c, sum: Integer;
begin
  SetLength(g, 3, 3);
  for r := 0 to High(g) do
    for c := 0 to High(g[r]) do
      g[r, c] := r + c;

  sum := 0;
  for r := 0 to High(g) do
    for c := 0 to High(g[r]) do
      sum := sum + g[r, c];

  WriteLn(sum);
end.
"#,
    );
    assert_eq!(out, vec!["18"]);
}

#[test]
fn test_dynarray_2d_clear_nil() {
    let out = run_pascal(
        r#"
program Test;
type TMatrix = array of array of Integer;
var m: TMatrix;
begin
  SetLength(m, 4, 4);
  m := nil;
  WriteLn(Length(m));
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_dynarray_2d_resize_preserve_contents() {
    let out = run_pascal(
        r#"
program Test;
type TMatrix = array of array of Integer;
var m: TMatrix;
begin
  SetLength(m, 2, 2);
  m[0, 0] := 5; m[0, 1] := 10;
  SetLength(m, 3, 3);
  WriteLn(m[0, 0]);
  WriteLn(m[0, 1]);
  WriteLn(Length(m));
end.
"#,
    );
    assert_eq!(out, vec!["5", "10", "3"]);
}

#[test]
fn test_dynarray_2d_procedure_var_param() {
    let out = run_pascal(
        r#"
program Test;
type TMatrix = array of array of Integer;
procedure InitIdentity(var m: TMatrix; size: Integer);
var i: Integer;
begin
  SetLength(m, size, size);
  for i := 0 to size - 1 do m[i, i] := 1;
end;
var mat: TMatrix;
begin
  InitIdentity(mat, 3);
  WriteLn(mat[0, 0].ToString + ',' + mat[1, 1].ToString + ',' + mat[2, 2].ToString);
  WriteLn(mat[0, 1]);
end.
"#,
    );
    assert_eq!(out, vec!["1,1,1", "0"]);
}

#[test]
fn test_dynarray_2d_string_matrix() {
    let out = run_pascal(
        r#"
program Test;
type TStrMatrix = array of array of String;
var sm: TStrMatrix;
begin
  SetLength(sm, 2, 2);
  sm[0, 0] := 'Alpha'; sm[0, 1] := 'Beta';
  sm[1, 0] := 'Gamma'; sm[1, 1] := 'Delta';
  WriteLn(sm[0, 0] + '-' + sm[1, 1]);
end.
"#,
    );
    assert_eq!(out, vec!["Alpha-Delta"]);
}

#[test]
fn test_dynarray_2d_record_matrix() {
    let out = run_pascal(
        r#"
program Test;
type TPoint = record X, Y: Integer; end;
type TPointMatrix = array of array of TPoint;
var pm: TPointMatrix;
begin
  SetLength(pm, 1, 1);
  pm[0, 0].X := 12; pm[0, 0].Y := 34;
  WriteLn(pm[0, 0].X.ToString + ':' + pm[0, 0].Y.ToString);
end.
"#,
    );
    assert_eq!(out, vec!["12:34"]);
}

#[test]
fn test_dynarray_2d_for_in_loop() {
    let out = run_pascal(
        r#"
program Test;
type TMatrix = array of array of Integer;
var m: TMatrix; row: array of Integer; elem: Integer;
begin
  SetLength(m, 2, 2);
  m[0, 0] := 1; m[0, 1] := 2;
  m[1, 0] := 3; m[1, 1] := 4;

  for row in m do
    for elem in row do
      WriteLn(elem);
end.
"#,
    );
    assert_eq!(out, vec!["1", "2", "3", "4"]);
}

#[test]
fn test_dynarray_2d_concat_rows() {
    let out = run_pascal(
        r#"
program Test;
type TRow = array of Integer;
type TMatrix = array of TRow;
var m: TMatrix; r1, r2: TRow;
begin
  r1 := [10, 20];
  r2 := [30, 40];
  m := [r1, r2];
  WriteLn(m[0][0]);
  WriteLn(m[1][1]);
end.
"#,
    );
    assert_eq!(out, vec!["10", "40"]);
}

#[test]
fn test_dynarray_2d_boolean_matrix() {
    let out = run_pascal(
        r#"
program Test;
type TBoolMatrix = array of array of Boolean;
var bm: TBoolMatrix;
begin
  SetLength(bm, 2, 2);
  bm[0, 0] := True; bm[1, 1] := True;
  WriteLn(bm[0, 0]);
  WriteLn(bm[0, 1]);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_dynarray_2d_transpose() {
    let out = run_pascal(
        r#"
program Test;
type TMatrix = array of array of Integer;
function Transpose(const src: TMatrix): TMatrix;
var r, c: Integer;
begin
  SetLength(Result, Length(src[0]), Length(src));
  for r := 0 to High(src) do
    for c := 0 to High(src[r]) do
      Result[c, r] := src[r, c];
end;

var m, t: TMatrix;
begin
  SetLength(m, 2, 3);
  m[0, 0] := 1; m[0, 1] := 2; m[0, 2] := 3;
  m[1, 0] := 4; m[1, 1] := 5; m[1, 2] := 6;

  t := Transpose(m);
  WriteLn(Length(t).ToString + 'x' + Length(t[0]).ToString);
  WriteLn(t[2, 0]); // m[0, 2] = 3
end.
"#,
    );
    assert_eq!(out, vec!["3x2", "3"]);
}

#[test]
fn test_dynarray_2d_empty_rows() {
    let out = run_pascal(
        r#"
program Test;
type TMatrix = array of array of Integer;
var m: TMatrix;
begin
  SetLength(m, 0, 0);
  WriteLn(Length(m) = 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_dynarray_2d_fillchar_bzero() {
    let out = run_pascal(
        r#"
program Test;
type TRow = array[0..1] of Byte;
type TMatrix = array of TRow;
var m: TMatrix;
begin
  SetLength(m, 2);
  FillChar(m[0][0], SizeOf(TRow), $AB);
  WriteLn(HexStr(m[0][0], 2));
end.
"#,
    );
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn test_dynarray_3d_ragged_cube() {
    let out = run_pascal(
        r#"
program Test;
type TRagged3D = array of array of array of Integer;
var r: TRagged3D;
begin
  SetLength(r, 1);
  SetLength(r[0], 2);
  SetLength(r[0, 0], 3);
  r[0, 0, 2] := 777;
  WriteLn(r[0, 0, 2]);
end.
"#,
    );
    assert_eq!(out, vec!["777"]);
}
