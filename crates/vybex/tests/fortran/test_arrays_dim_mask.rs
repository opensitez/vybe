use super::helpers::compile_ok;

// ── SUM with DIM ──────────────────────────────────────────────

#[test] fn sum_dim1() {
    compile_ok(r#"
program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: col_sums(4)
    col_sums = sum(m, dim=1)
    print *, col_sums(1)
end program test
"#);
}

#[test] fn sum_dim2() {
    compile_ok(r#"
program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: row_sums(3)
    row_sums = sum(m, dim=2)
    print *, row_sums(1)
end program test
"#);
}

#[test] fn sum_with_mask() {
    compile_ok(r#"
program test
    integer :: a(6) = [1, 2, 3, 4, 5, 6]
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    print *, sum(a, mask=mask)
end program test
"#);
}

#[test] fn sum_dim1_with_mask() {
    compile_ok(r#"
program test
    integer :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])
    logical :: mask(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])
    integer :: col_sums(3)
    col_sums = sum(m, dim=1, mask=mask)
    print *, col_sums(1)
end program test
"#);
}

// ── PRODUCT with DIM / MASK ───────────────────────────────────

#[test] fn product_dim1() {
    compile_ok(r#"
program test
    integer :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])
    integer :: col_prod(3)
    col_prod = product(m, dim=1)
    print *, col_prod(1)
end program test
"#);
}

#[test] fn product_with_mask() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    logical :: mask(5) = [.true., .true., .false., .true., .false.]
    print *, product(a, mask=mask)
end program test
"#);
}

// ── MAXVAL / MINVAL with DIM / MASK ──────────────────────────

#[test] fn maxval_dim1() {
    compile_ok(r#"
program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: col_max(4)
    col_max = maxval(m, dim=1)
    print *, col_max(1)
end program test
"#);
}

#[test] fn maxval_dim2() {
    compile_ok(r#"
program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: row_max(3)
    row_max = maxval(m, dim=2)
    print *, row_max(1)
end program test
"#);
}

#[test] fn maxval_with_mask() {
    compile_ok(r#"
program test
    integer :: a(6) = [1, 9, 2, 8, 3, 7]
    logical :: mask(6) = [.false., .false., .true., .true., .true., .true.]
    print *, maxval(a, mask=mask)
end program test
"#);
}

#[test] fn minval_dim1() {
    compile_ok(r#"
program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: col_min(4)
    col_min = minval(m, dim=1)
    print *, col_min(2)
end program test
"#);
}

#[test] fn minval_with_mask() {
    compile_ok(r#"
program test
    integer :: a(6) = [10, 1, 20, 2, 30, 3]
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    print *, minval(a, mask=mask)
end program test
"#);
}

// ── ALL / ANY with DIM / MASK ─────────────────────────────────

#[test] fn all_dim1() {
    compile_ok(r#"
program test
    logical :: m(2,3) = reshape([.true.,.true.,.true.,.false.,.true.,.true.],[2,3])
    logical :: col_all(3)
    col_all = all(m, dim=1)
    print *, col_all(1)
    print *, col_all(2)
end program test
"#);
}

#[test] fn all_dim2() {
    compile_ok(r#"
program test
    logical :: m(3,2) = reshape([.true.,.true.,.true.,.true.,.false.,.true.],[3,2])
    logical :: row_all(3)
    row_all = all(m, dim=2)
    print *, row_all(1)
end program test
"#);
}

#[test] fn any_dim1() {
    compile_ok(r#"
program test
    logical :: m(2,3) = reshape([.false.,.false.,.true.,.false.,.false.,.false.],[2,3])
    logical :: col_any(3)
    col_any = any(m, dim=1)
    print *, col_any(1)
    print *, col_any(2)
end program test
"#);
}

#[test] fn any_dim2() {
    compile_ok(r#"
program test
    logical :: m(3,2) = reshape([.false.,.true.,.false.,.false.,.false.,.false.],[3,2])
    logical :: row_any(3)
    row_any = any(m, dim=2)
    print *, row_any(1)
end program test
"#);
}

// ── COUNT with DIM / MASK ─────────────────────────────────────

#[test] fn count_basic_mask() {
    compile_ok(r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    print *, count(a > 3)
end program test
"#);
}

#[test] fn count_dim1() {
    compile_ok(r#"
program test
    integer :: m(3,4) = reshape([(i, i=1,12)],[3,4])
    integer :: col_count(4)
    col_count = count(m > 6, dim=1)
    print *, col_count(1)
    print *, col_count(3)
end program test
"#);
}

#[test] fn count_dim2() {
    compile_ok(r#"
program test
    logical :: m(2,3) = reshape([.true.,.false.,.true.,.true.,.false.,.true.],[2,3])
    integer :: row_count(2)
    row_count = count(m, dim=2)
    print *, row_count(1)
end program test
"#);
}

// ── MAXLOC / MINLOC with DIM / MASK ──────────────────────────

#[test] fn maxloc_with_mask() {
    compile_ok(r#"
program test
    integer :: a(6) = [1, 9, 2, 8, 3, 7]
    logical :: mask(6) = [.false., .false., .true., .true., .true., .true.]
    integer :: loc(1)
    loc = maxloc(a, mask=mask)
    print *, loc(1)
end program test
"#);
}

#[test] fn minloc_with_mask() {
    compile_ok(r#"
program test
    integer :: a(5) = [5, 1, 4, 1, 5]
    logical :: mask(5) = [.true., .false., .true., .true., .true.]
    integer :: loc(1)
    loc = minloc(a, mask=mask)
    print *, loc(1)
end program test
"#);
}

#[test] fn maxloc_dim1() {
    compile_ok(r#"
program test
    integer :: m(3,3) = reshape([1,9,2,8,3,7,4,6,5],[3,3])
    integer :: col_maxloc(3)
    col_maxloc = maxloc(m, dim=1)
    print *, col_maxloc(1)
end program test
"#);
}

#[test] fn minloc_dim2() {
    compile_ok(r#"
program test
    integer :: m(3,3) = reshape([1,9,2,8,3,7,4,6,5],[3,3])
    integer :: row_minloc(3)
    row_minloc = minloc(m, dim=2)
    print *, row_minloc(1)
end program test
"#);
}

// ── FINDLOC with DIM / MASK ───────────────────────────────────

#[test] fn findloc_with_mask() {
    compile_ok(r#"
program test
    integer :: a(6) = [1, 2, 1, 2, 1, 2]
    logical :: mask(6) = [.false., .true., .true., .true., .true., .true.]
    integer :: loc(1)
    loc = findloc(a, 1, mask=mask)
    print *, loc(1)
end program test
"#);
}

#[test] fn findloc_dim() {
    compile_ok(r#"
program test
    integer :: m(3,3) = reshape([1,2,1,2,1,2,1,2,1],[3,3])
    integer :: col_loc(3)
    col_loc = findloc(m, 2, dim=1)
    print *, col_loc(1)
end program test
"#);
}

// ── SIZE with DIM ─────────────────────────────────────────────

#[test] fn size_dim1() {
    compile_ok(r#"
program test
    integer :: m(3,4,5)
    print *, size(m, 1)
    print *, size(m, 2)
    print *, size(m, 3)
end program test
"#);
}

#[test] fn size_kind_param() {
    compile_ok(r#"
program test
    use iso_fortran_env
    integer :: a(100)
    integer(int64) :: n
    n = size(a, kind=int64)
    print *, n
end program test
"#);
}

// ── SHAPE with KIND ───────────────────────────────────────────

#[test] fn shape_with_kind() {
    compile_ok(r#"
program test
    use iso_fortran_env
    real :: m(3,4)
    integer(int64), allocatable :: sh(:)
    sh = shape(m, kind=int64)
    print *, sh(1), sh(2)
end program test
"#);
}

// ── Combining DIM and MASK ────────────────────────────────────

#[test] fn sum_dim_mask_combined() {
    compile_ok(r#"
program test
    integer :: m(4,4) = reshape([(i, i=1,16)],[4,4])
    logical :: mask(4,4)
    integer :: row_sums(4)
    mask = m > 8
    row_sums = sum(m, dim=2, mask=mask)
    print *, row_sums(1)
end program test
"#);
}
