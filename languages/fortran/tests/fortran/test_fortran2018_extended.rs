//! Extended Fortran 2018 coverage: REDUCE variants, SORT, SIZE/LBOUND/UBOUND/SHAPE
//! with KIND=, OUT_OF_RANGE edge cases, TYPEOF, RANDOM_INIT combinations, and deeper
//! assumed-rank SELECT RANK forms. Distinct from `test_fortran2018.rs`.

use super::helpers::{compile_ok, run_prints};

fortran_cases! {
    // ── REDUCE extended (not in test_fortran2018.rs) ─────────────────

    reduce_builtin_min_operator => {
        "program t\ninteger :: a(4) = [3, 1, 4, 2]\nprint *, reduce(a, operator(min))\nend program t\n",
        ["1"]
    };

    reduce_logical_and_chain => {
        "program t\nlogical :: flags(3) = [.true., .true., .false.]\nprint *, reduce(flags, operator(.and.))\nend program t\n",
        ["false"]
    };

    reduce_logical_or_chain => {
        "program t\nlogical :: flags(3) = [.false., .true., .false.]\nprint *, reduce(flags, operator(.or.))\nend program t\n",
        ["true"]
    };

    reduce_real_sum_three_values => {
        "program t\nreal :: vals(3) = [1.5, 2.5, 3.0]\nprint *, reduce(vals, operator(+))\nend program t\n",
        ["7"]
    };

    reduce_product_with_identity_one => {
        "program t\ninteger :: a(4) = [2, 3, 4, 5]\nprint *, reduce(a, operator(*), identity=1)\nend program t\n",
        ["120"]
    };

    reduce_dim1_with_mask_columns => {
        "program t\ninteger :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])\nlogical :: mask(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])\ninteger :: r(3)\nr = reduce(m, operator(+), dim=1, mask=mask)\nprint *, r(1)\nprint *, r(2)\nprint *, r(3)\nend program t\n",
        ["1", "5", "3"]
    };

    reduce_custom_min_procedure => {
        "program t\ninteger :: a(5) = [8, 3, 9, 1, 6]\nprint *, reduce(a, pick_min)\ncontains\npure function pick_min(x, y) result(r)\ninteger, intent(in) :: x, y\ninteger :: r\nr = min(x, y)\nend function pick_min\nend program t\n",
        ["1"]
    };

    // ── SIZE / LBOUND / UBOUND / SHAPE with KIND= (F2018) ───────────

    size_matrix_dim2_with_int64_kind => {
        "program t\nuse iso_fortran_env\ninteger :: m(3,4)\nprint *, size(m, 2, kind=int64)\nend program t\n",
        ["4"]
    };

    lbound_rank1_with_int64_kind => {
        "program t\nuse iso_fortran_env\ninteger :: a(5)\nprint *, lbound(a, 1, kind=int64)\nend program t\n",
        ["1"]
    };

    ubound_nondefault_lower_with_int64_kind => {
        "program t\nuse iso_fortran_env\ninteger :: a(2:6)\nprint *, ubound(a, 1, kind=int64)\nend program t\n",
        ["6"]
    };

    shape_matrix_with_int64_kind_elements => {
        "program t\nuse iso_fortran_env\ninteger :: m(2,5)\ninteger(int64) :: sh(2)\nsh = shape(m, kind=int64)\nprint *, sh(1)\nprint *, sh(2)\nend program t\n",
        ["2", "5"]
    };

    // ── OUT_OF_RANGE extended boolean cases ─────────────────────────

    out_of_range_fifty_fits_int8 => {
        "program t\ninteger :: x = 50\nprint *, out_of_range(x, 0_1)\nend program t\n",
        ["false"]
    };

    out_of_range_two_hundred_exceeds_int8 => {
        "program t\ninteger :: x = 200\nprint *, out_of_range(x, 0_1)\nend program t\n",
        ["true"]
    };

    out_of_range_zero_fits_int16 => {
        "program t\ninteger :: x = 0\nprint *, out_of_range(x, 0_2)\nend program t\n",
        ["false"]
    };
}

// ── SORT intrinsic (Fortran 2018) ────────────────────────────────

#[test]
fn sort_integer_vector_ascending() {
    compile_ok(
        r#"
program t
    integer :: a(5) = [3, 1, 4, 1, 5]
    call sort(a)
    print *, a(1), a(5)
end program t
"#,
    );
}

#[test]
fn sort_integer_vector_descending() {
    compile_ok(
        r#"
program t
    integer :: a(4) = [3, 1, 4, 2]
    call sort(a, reverse=.true.)
    print *, a(1)
end program t
"#,
    );
}

#[test]
fn sort_matrix_along_dim1() {
    compile_ok(
        r#"
program t
    integer :: m(2,3) = reshape([3, 1, 4, 1, 5, 9], [2, 3])
    call sort(m, dim=1)
    print *, m(1, 1), m(2, 1)
end program t
"#,
    );
}

#[test]
fn sort_matrix_with_mask() {
    compile_ok(
        r#"
program t
    integer :: a(6) = [5, 2, 8, 1, 9, 3]
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    call sort(a, mask=mask)
    print *, a(1)
end program t
"#,
    );
}

// ── TYPEOF intrinsic (Fortran 2018) ──────────────────────────────

#[test]
fn typeof_integer_scalar() {
    compile_ok(
        r#"
program t
    integer :: x = 7
    print *, typeof(x)
end program t
"#,
    );
}

#[test]
fn typeof_real_scalar() {
    compile_ok(
        r#"
program t
    real :: x = 2.5
    print *, typeof(x)
end program t
"#,
    );
}

#[test]
fn typeof_logical_scalar() {
    compile_ok(
        r#"
program t
    logical :: flag = .true.
    print *, typeof(flag)
end program t
"#,
    );
}

#[test]
fn typeof_integer_vector() {
    compile_ok(
        r#"
program t
    integer :: v(3) = [1, 2, 3]
    print *, typeof(v)
end program t
"#,
    );
}

// ── OUT_OF_RANGE compile-only edge cases ───────────────────────

#[test]
fn out_of_range_quiet_nan_to_integer() {
    compile_ok(
        r#"
program t
    use ieee_arithmetic
    real :: x
    x = ieee_value(x, ieee_quiet_nan)
    print *, out_of_range(x, 0)
end program t
"#,
    );
}

#[test]
fn out_of_range_round_false_on_boundary() {
    compile_ok(
        r#"
program t
    real :: x = 127.6
    print *, out_of_range(x, 0_1, round=.false.)
end program t
"#,
    );
}

#[test]
fn out_of_range_integer_to_smaller_kind_negative() {
    compile_ok(
        r#"
program t
    integer :: x = -200
    print *, out_of_range(x, 0_1)
end program t
"#,
    );
}

// ── RANDOM_INIT combinations ───────────────────────────────────

#[test]
fn random_init_repeatable_and_image_distinct() {
    compile_ok(
        r#"
program t
    call random_init(repeatable=.true., image_distinct=.true.)
    real :: r
    call random_number(r)
    print *, r >= 0.0
end program t
"#,
    );
}

#[test]
fn random_init_in_module_initializer() {
    compile_ok(
        r#"
module rng
    implicit none
contains
    subroutine seed_once()
        call random_init(repeatable=.false., image_distinct=.false.)
    end subroutine seed_once
end module rng

program t
    use rng
    call seed_once()
    print *, 'seeded'
end program t
"#,
    );
}

// ── Assumed-rank / SELECT RANK depth ────────────────────────────

#[test]
fn select_rank_explicit_rank3_branch() {
    compile_ok(
        r#"
program t
    call inspect(reshape([(i, i = 1, 24)], [2, 3, 4]))
contains
    subroutine inspect(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(3)
            print *, size(x, 1), size(x, 2), size(x, 3)
        rank default
            print *, rank(x)
        end select
    end subroutine inspect
end program t
"#,
    );
}

#[test]
fn assumed_rank_module_procedure_rank2() {
    compile_ok(
        r#"
module ranks
    implicit none
contains
    subroutine rows(x)
        real, intent(in) :: x(..)
        select rank(x)
        rank(2)
            print *, size(x, 1)
        rank default
            print *, 0
        end select
    end subroutine rows
end module ranks

program t
    use ranks
    real :: grid(4, 3)
    call rows(grid)
end program t
"#,
    );
}

#[test]
fn select_rank_catches_rank1_vector() {
    let out = run_prints(
        r#"
program t
    call tag([10, 20, 30])
contains
    subroutine tag(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(1)
            print *, size(x)
        rank default
            print *, 0
        end select
    end subroutine tag
end program t
"#,
    );
    assert_eq!(out, ["3"]);
}

// ── IMPLICIT NONE (type, external) in module scope ──────────────

#[test]
fn implicit_none_type_external_in_module() {
    compile_ok(
        r#"
module guarded
    implicit none (type, external)
contains
    function twice(n) result(r)
        integer, intent(in) :: n
        integer :: r
        r = n * 2
    end function twice
end module guarded

program t
    use guarded
    print *, twice(11)
end program t
"#,
    );
}

// ── ERROR STOP with character (Fortran 2018) ────────────────────

#[test]
fn error_stop_character_message() {
    compile_ok(
        r#"
program t
    logical :: ok = .true.
    if (.not. ok) error stop 'aborted'
    print *, 'fine'
end program t
"#,
    );
}

// ── REDUCE ordered + mask combined (compile-only) ───────────────

#[test]
fn reduce_ordered_with_mask() {
    compile_ok(
        r#"
program t
    integer :: a(6) = [6, 5, 4, 3, 2, 1]
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    integer :: r
    r = reduce(a, operator(+), mask=mask, ordered=.true.)
    print *, r
end program t
"#,
    );
}
