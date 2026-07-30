use super::helpers::{compile_ok, run_prints};

// ── WHERE with ELSEWHERE ──────────────────────────────────────

#[test]
fn where_else_basic() {
    let out = run_prints(
        r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    where (a > 3)
        b = a * 10
    elsewhere
        b = a
    end where
    print *, b(1)
    print *, b(5)
end program test
"#,
    );

    assert_eq!(out, vec!["1", "50"]);
}

#[test]
fn where_else_basic_runtime() {
    let out = run_prints(
        r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(5)
    where (a > 3)
        b = a * 10
    elsewhere
        b = a
    end where
    print *, b(1)
    print *, b(5)
end program test
"#,
    );
    assert_eq!(out, vec!["1", "50"]);
}

#[test]
fn where_else_real_runtime() {
    let out = run_prints(
        r#"
program test
    real :: x(6) = [-2., -1., 0., 1., 2., 3.]
    real :: y(6)
    where (x >= 0.0)
        y = sqrt(x)
    elsewhere
        y = 0.0
    end where
    print *, nint(y(1)*1000)
    print *, nint(y(4)*1000)
end program test
"#,
    );
    assert_eq!(out, vec!["0", "1000"]);
}

#[test]
fn where_multi_elsewhere_runtime() {
    let out = run_prints(
        r#"
program test
    integer :: a(6) = [1, 5, 10, 50, 100, 500]
    character(len=6) :: b(6)
    where (a < 10)
        b = 'small '
    elsewhere (a < 100)
        b = 'medium'
    elsewhere
        b = 'large '
    end where
    print *, trim(b(1))
    print *, trim(b(3))
    print *, trim(b(5))
end program test
"#,
    );
    assert_eq!(out, vec!["small", "medium", "large"]);
}

#[test]
fn where_2d_basic_runtime() {
    let out = run_prints(
        r#"
program test
    integer :: m(3,3) = reshape([1,2,3,4,5,6,7,8,9],[3,3])
    where (m > 5)
        m = m * 2
    end where
    print *, m(1,1)
    print *, m(3,3)
end program test
"#,
    );
    assert_eq!(out, vec!["1", "18"]);
}

#[test]
fn where_mask_with_mod_runtime() {
    let out = run_prints(
        r#"
program test
    integer :: a(10) = [(i, i=1,10)]
    integer :: b(10)
    b = 0
    where (mod(a, 3) == 0)
        b = a
    end where
    print *, b(3)
    print *, b(6)
    print *, b(1)
end program test
"#,
    );
    assert_eq!(out, vec!["3", "6", "0"]);
}

#[test]
fn where_in_subroutine_runtime() {
    let out = run_prints(
        r#"
program test
    integer :: a(5) = [3, -1, 5, -2, 4]
    call clamp_negatives(a)
    print *, a(2)
    print *, a(4)
    print *, a(1)
contains
    subroutine clamp_negatives(x)
        integer, intent(inout) :: x(:)
        where (x < 0)
            x = 0
        end where
    end subroutine clamp_negatives
end program test
"#,
    );
    assert_eq!(out, vec!["0", "0", "3"]);
}

#[test]
fn where_in_module_function_runtime() {
    let out = run_prints(
        r#"
module where_mod
    implicit none
contains
    function positive_part(a) result(b)
        real, intent(in) :: a(:)
        real :: b(size(a))
        b = 0.0
        where (a > 0.0)
            b = a
        end where
    end function positive_part
end module where_mod

program test
    use where_mod
    real :: v(5) = [-1., 2., -3., 4., -5.]
    real :: p(5)
    p = positive_part(v)
    print *, nint(p(2))
    print *, nint(p(1))
    print *, nint(p(4))
end program test
"#,
    );
    assert_eq!(out, vec!["2", "0", "4"]);
}

#[test]
fn where_else_real() {
    let out = run_prints(
        r#"
program test
    real :: x(6) = [-2., -1., 0., 1., 2., 3.]
    real :: y(6)
where (x >= 0.0)
        y = sqrt(x)
    elsewhere
        y = 0.0
    end where
    print *, nint(y(1))
    print *, nint(y(4))
end program test
"#,
    );

    assert_eq!(out, vec!["0", "1"]);
}

#[test]
fn where_else_set_zero() {
    compile_ok(
        r#"
program test
    integer :: a(6) = [10, -2, 5, -8, 3, -1]
    where (a < 0)
        a = 0
    elsewhere
        a = a
    end where
    print *, a(2)
    print *, a(1)
end program test
"#,
    );
}

// ── Multiple ELSEWHERE clauses ────────────────────────────────

#[test]
fn where_multi_elsewhere() {
    let out = run_prints(
        r#"
program test
    integer :: a(6) = [1, 5, 10, 50, 100, 500]
    character(len=6) :: b(6)
    where (a < 10)
        b = 'small '
    elsewhere (a < 100)
        b = 'medium'
    elsewhere
        b = 'large '
    end where
    print *, trim(b(1))
    print *, trim(b(3))
    print *, trim(b(5))
end program test
"#,
    );

    assert_eq!(out, vec!["small", "medium", "large"]);
}

#[test]
fn where_three_elsewhere_clauses() {
    compile_ok(
        r#"
program test
    real :: t(6) = [-10., -1., 0., 1., 10., 100.]
    integer :: cat(6)
    where (t < -5.0)
        cat = 1
    elsewhere (t < 0.0)
        cat = 2
    elsewhere (t < 5.0)
        cat = 3
    elsewhere
        cat = 4
    end where
    print *, cat(1)
    print *, cat(2)
    print *, cat(4)
    print *, cat(6)
end program test
"#,
    );
}

#[test]
fn where_multi_elsewhere_order_runtime() {
    let out = run_prints(
        r#"
program test
    integer :: a(6) = [1, 5, 10, 50, 100, 500]
    integer :: b(6)
    where (a < 10)
        b = 1
    elsewhere (a == 100)
        b = 2
    elsewhere
        b = 3
    end where
    print *, b(1)
    print *, b(3)
    print *, b(5)
    print *, b(6)
end program test
"#,
    );
    assert_eq!(out, vec!["1", "1", "3", "3"]);
}

#[test]
fn where_without_else_masks_only_true_elements() {
    let out = run_prints(
        r#"
program test
    integer :: a(4) = [1, 2, 3, 4]
    where (a >= 3)
        a = a * 10
    end where
    print *, a(1)
    print *, a(3)
    print *, a(4)
end program test
"#,
    );
    assert_eq!(out, vec!["1", "30", "40"]);
}

#[test]
fn where_nested_elsewhere_runtime_layers() {
    let out = run_prints(
        r#"
program test
    integer :: a(6) = [1, 10, 2, 20, 3, 30]
    integer :: b(6) = 0
    where (a > 5)
        where (a > 15)
            b = a * 100
        elsewhere
            b = a * 10
        end where
    end where
    print *, b(1)
    print *, b(2)
    print *, b(4)
end program test
"#,
    );
    assert_eq!(out, vec!["0", "100", "2000"]);
}

// ── WHERE on 2D arrays ────────────────────────────────────────

#[test]
fn where_2d_basic() {
    let out = run_prints(
        r#"
program test
    integer :: m(3,3) = reshape([1,2,3,4,5,6,7,8,9],[3,3])
    where (m > 5)
        m = m * 2
    end where
    print *, m(1,1)
    print *, m(3,3)
end program test
"#,
    );

    assert_eq!(out, vec!["1", "18"]);
}

#[test]
fn where_2d_else() {
    let out = run_prints(
        r#"
program test
    real :: m(4,4) = reshape([(real(i), i=1,16)],[4,4])
    real :: result(4,4)
where (m > 8.0)
        result = m
    elsewhere
        result = 0.0
    end where
    print *, nint(result(1,1))
    print *, nint(result(4,4))
end program test
"#,
    );

    assert_eq!(out, vec!["0", "16"]);
}

// ── WHERE with function call in mask ─────────────────────────

#[test]
fn where_mask_with_mod() {
    let out = run_prints(
        r#"
program test
    integer :: a(10) = [(i, i=1,10)]
    integer :: b(10)
    b = 0
    where (mod(a, 3) == 0)
        b = a
    end where
    print *, b(3)
    print *, b(6)
    print *, b(1)
end program test
"#,
    );

    assert_eq!(out, vec!["3", "6", "0"]);
}

#[test]
fn where_mask_with_abs() {
    let out = run_prints(
        r#"
program test
    real :: a(6) = [-3., 2., -1., 4., -5., 0.]
    real :: b(6)
    b = 0.0
where (abs(a) > 2.0)
        b = a
    end where
    print *, nint(b(1))
    print *, nint(b(2))
end program test
"#,
    );

    assert_eq!(out, vec!["-3", "0"]);
}

// ── Nested WHERE ──────────────────────────────────────────────

#[test]
fn nested_where() {
    let out = run_prints(
        r#"
program test
    integer :: a(6) = [1, 10, 2, 20, 3, 30]
    integer :: b(6) = 0
    where (a > 5)
        where (a > 15)
            b = a * 100
        elsewhere
            b = a * 10
        end where
    elsewhere
        b = a
    end where
    print *, b(1)
    print *, b(2)
    print *, b(4)
end program test
"#,
    );

    assert_eq!(out, vec!["1", "100", "2000"]);
}

#[test]
fn where_all_false_mask_keeps_input_array_unmodified() {
    let out = run_prints(
        r#"
program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: b(4)
    b = 0
    where (a < 0)
        b = 9
    end where
    print *, b(1)
    print *, b(2)
    print *, b(3)
    print *, b(4)
end program test
"#,
    );

    assert_eq!(out, vec!["0", "0", "0", "0"]);
}

#[test]
fn where_elsewhere_chain_falls_to_default_for_unmatched_mask() {
    let out = run_prints(
        r#"
program test
    integer :: a(4) = [5, 15, 25, 35]
    character(len=6) :: b(4)
    where (a < 10)
        b = 'low'
    elsewhere (a < 20)
        b = 'mid'
    elsewhere (a < 30)
        b = 'high'
    elsewhere
        b = 'top'
    end where
    print *, trim(b(1))
    print *, trim(b(2))
    print *, trim(b(3))
    print *, trim(b(4))
end program test
"#,
    );

    assert_eq!(out, vec!["low", "mid", "high", "top"]);
}

#[test]
fn nested_where_2d() {
    compile_ok(
        r#"
program test
    real :: m(4,4) = reshape([(real(i-8), i=1,16)],[4,4])
    real :: result(4,4)
    result = 0.0
    where (m > 0.0)
        where (m > 4.0)
            result = m * 2.0
        elsewhere
            result = m
        end where
    end where
    print *, result(1,1)
end program test
"#,
    );
}

// ── WHERE in subroutine ───────────────────────────────────────

#[test]
fn where_in_subroutine() {
    compile_ok(
        r#"
program test
    integer :: a(5) = [3, -1, 5, -2, 4]
    call clamp_negatives(a)
    print *, a(2)
    print *, a(4)
contains
    subroutine clamp_negatives(x)
        integer, intent(inout) :: x(:)
        where (x < 0)
            x = 0
        end where
    end subroutine clamp_negatives
end program test
"#,
    );
}

#[test]
fn where_in_module_function() {
    compile_ok(
        r#"
module where_mod
    implicit none
contains
    function positive_part(a) result(b)
        real, intent(in) :: a(:)
        real :: b(size(a))
        b = 0.0
        where (a > 0.0)
            b = a
        end where
    end function positive_part
end module where_mod

program test
    use where_mod
    real :: v(5) = [-1., 2., -3., 4., -5.]
    real :: p(5)
    p = positive_part(v)
    print *, p(2)
    print *, p(1)
end program test
"#,
    );
}

// ── STORAGE_SIZE intrinsic ────────────────────────────────────

#[test]
fn storage_size_integer() {
    compile_ok(
        r#"
program test
    integer :: x = 0
    print *, storage_size(x)
end program test
"#,
    );
}

#[test]
fn storage_size_real() {
    compile_ok(
        r#"
program test
    real :: x = 0.0
    print *, storage_size(x)
end program test
"#,
    );
}

#[test]
fn storage_size_double() {
    compile_ok(
        r#"
program test
    real(kind=8) :: x = 0.0d0
    print *, storage_size(x)
end program test
"#,
    );
}

#[test]
fn storage_size_runtime_basics() {
    let out = run_prints(
        r#"
program test
    integer :: i = 0
    real :: r = 0.0
    real(kind=8) :: d = 0.0d0
    print *, storage_size(i)
    print *, storage_size(r)
    print *, storage_size(d)
end program test
"#,
    );
    assert_eq!(out, vec!["32", "32", "64"]);
}

#[test]
fn storage_size_logical() {
    compile_ok(
        r#"
program test
    logical :: b = .false.
    print *, storage_size(b)
end program test
"#,
    );
}

#[test]
fn storage_size_complex() {
    compile_ok(
        r#"
program test
    complex :: c = (0., 0.)
    real :: r = 0.
    print *, storage_size(c) == 2 * storage_size(r)
end program test
"#,
    );
}

#[test]
fn storage_size_derived_type() {
    compile_ok(
        r#"
program test
    type :: Pair
        integer :: x, y
    end type Pair
    type(Pair) :: p
    print *, storage_size(p)
end program test
"#,
    );
}

#[test]
fn storage_size_int8() {
    compile_ok(
        r#"
program test
    integer(kind=8) :: big = 0_8
    integer(kind=4) :: small = 0_4
    print *, storage_size(big) > storage_size(small)
end program test
"#,
    );
}

#[test]
fn storage_size_with_kind() {
    compile_ok(
        r#"
program test
    use iso_fortran_env
    integer :: x = 0
    integer(int64) :: n
    n = storage_size(x, kind=int64)
    print *, n
end program test
"#,
    );
}

// ── BIT_SIZE intrinsic ────────────────────────────────────────

#[test]
fn bit_size_int4() {
    compile_ok(
        r#"
program test
    integer :: x = 0
    print *, bit_size(x)
end program test
"#,
    );
}

#[test]
fn bit_size_int8() {
    compile_ok(
        r#"
program test
    integer(kind=8) :: x = 0_8
    print *, bit_size(x)
end program test
"#,
    );
}

#[test]
fn bit_size_int2() {
    compile_ok(
        r#"
program test
    integer(kind=2) :: x = 0_2
    print *, bit_size(x)
end program test
"#,
    );
}

#[test]
fn bit_size_int1() {
    compile_ok(
        r#"
program test
    integer(kind=1) :: x = 0_1
    print *, bit_size(x)
end program test
"#,
    );
}

#[test]
fn bit_size_vs_storage_size() {
    compile_ok(
        r#"
program test
    integer :: x = 0
    print *, bit_size(x) == storage_size(x)
end program test
"#,
    );
}

#[test]
fn bit_size_array() {
    compile_ok(
        r#"
program test
    integer :: a(10)
    print *, bit_size(a(1))
end program test
"#,
    );
}

#[test]
fn bit_size_in_expression() {
    compile_ok(
        r#"
program test
    integer :: x = 1
    integer :: half_bits
    half_bits = bit_size(x) / 2
    print *, half_bits
end program test
"#,
    );
}
