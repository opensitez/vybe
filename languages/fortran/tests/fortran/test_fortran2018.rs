use super::helpers::{compile_ok, run_prints};

// ── REDUCE intrinsic (Fortran 2018) ──────────────────────────

#[test]
fn reduce_sum() {
    compile_ok(
        r#"
program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: total
    total = reduce(a, operator(+))
    print *, total
end program test
"#,
    );
}

#[test]
fn reduce_product() {
    compile_ok(
        r#"
program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: prod
    prod = reduce(a, operator(*))
    print *, prod
end program test
"#,
    );
}

#[test]
fn reduce_max() {
    compile_ok(
        r#"
program test
    integer :: a(5) = [3, 1, 4, 1, 5]
    integer :: m
    m = reduce(a, my_max)
    print *, m
contains
    pure function my_max(x, y) result(r)
        integer, intent(in) :: x, y
        integer :: r
        r = max(x, y)
    end function my_max
end program test
"#,
    );
}

#[test]
fn reduce_with_identity() {
    compile_ok(
        r#"
program test
    integer :: a(4) = [1, 2, 3, 4]
    integer :: total
    total = reduce(a, operator(+), identity=0)
    print *, total
end program test
"#,
    );
}

#[test]
fn reduce_ordered() {
    compile_ok(
        r#"
program test
    integer :: a(5) = [5, 4, 3, 2, 1]
    integer :: r
    r = reduce(a, operator(+), ordered=.true.)
    print *, r
end program test
"#,
    );
}

#[test]
fn reduce_dim() {
    compile_ok(
        r#"
program test
    integer :: m(3,3) = reshape([1,2,3,4,5,6,7,8,9],[3,3])
    integer :: row_sums(3)
    row_sums = reduce(m, operator(+), dim=2)
    print *, row_sums(1)
end program test
"#,
    );
}

#[test]
fn reduce_mask() {
    compile_ok(
        r#"
program test
    integer :: a(6) = [1, 2, 3, 4, 5, 6]
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    integer :: r
    r = reduce(a, operator(+), mask=mask)
    print *, r
end program test
"#,
    );
}

// ── OUT_OF_RANGE intrinsic (Fortran 2018) ─────────────────────

#[test]
fn out_of_range_integer_in_range() {
    compile_ok(
        r#"
program test
    integer(kind=4) :: x = 100
    print *, out_of_range(x, 0_2)
end program test
"#,
    );
}

#[test]
fn out_of_range_integer_overflow() {
    compile_ok(
        r#"
program test
    integer(kind=8) :: big = 1000000000000_8
    print *, out_of_range(big, 0_2)
end program test
"#,
    );
}

#[test]
fn out_of_range_real_to_integer() {
    compile_ok(
        r#"
program test
    real :: x = 3.14
    print *, out_of_range(x, 0)
end program test
"#,
    );
}

#[test]
fn out_of_range_real_infinity() {
    compile_ok(
        r#"
program test
    use ieee_arithmetic
    real :: x
    x = ieee_value(x, ieee_positive_inf)
    print *, out_of_range(x, 0)
end program test
"#,
    );
}

#[test]
fn out_of_range_real_to_real() {
    compile_ok(
        r#"
program test
    real(kind=8) :: d = 1.0d38
    print *, out_of_range(d, 0.0_4)
end program test
"#,
    );
}

#[test]
fn out_of_range_with_round() {
    compile_ok(
        r#"
program test
    real :: x = 127.6
    print *, out_of_range(x, 0_1, round=.true.)
end program test
"#,
    );
}

// ── RANDOM_INIT (Fortran 2018) ────────────────────────────────

#[test]
fn random_init_repeatable() {
    compile_ok(
        r#"
program test
    call random_init(repeatable=.true., image_distinct=.false.)
    real :: x
    call random_number(x)
    print *, x >= 0.0 .and. x < 1.0
end program test
"#,
    );
}

#[test]
fn random_init_non_repeatable() {
    compile_ok(
        r#"
program test
    call random_init(repeatable=.false., image_distinct=.true.)
    real :: r
    call random_number(r)
    print *, 'ok'
end program test
"#,
    );
}

// ── Assumed-rank arrays — dimension(..) (Fortran 2018) ────────

#[test]
fn assumed_rank_basic() {
    compile_ok(
        r#"
module ar_mod
    implicit none
contains
    subroutine describe(x)
        real, intent(in) :: x(..)
        print *, rank(x)
    end subroutine describe
end module ar_mod

program test
    use ar_mod
    real :: a(3,4)
    call describe(a)
end program test
"#,
    );
}

#[test]
fn assumed_rank_scalar() {
    compile_ok(
        r#"
module ar_mod
    implicit none
contains
    subroutine show_rank(x)
        integer, intent(in) :: x(..)
        print *, rank(x)
    end subroutine show_rank
end module ar_mod

program test
    use ar_mod
    integer :: s = 42
    call show_rank(s)
end program test
"#,
    );
}

#[test]
fn assumed_rank_with_select_rank() {
    compile_ok(
        r#"
module ar_mod
    implicit none
contains
    subroutine process(x)
        real, intent(in) :: x(..)
        select rank(x)
        rank(0)
            print *, 'scalar', x
        rank(1)
            print *, 'vector of size', size(x)
        rank(2)
            print *, 'matrix', size(x,1), 'x', size(x,2)
        rank default
            print *, 'rank', rank(x)
        end select
    end subroutine process
end module ar_mod

program test
    use ar_mod
    real :: v(5) = [1., 2., 3., 4., 5.]
    call process(v)
end program test
"#,
    );
}

#[test]
fn assumed_rank_2d() {
    compile_ok(
        r#"
module ar_mod
    implicit none
contains
    subroutine show(a)
        integer, intent(in) :: a(..)
        select rank(a)
        rank(2)
            print *, size(a,1), size(a,2)
        rank default
            print *, rank(a)
        end select
    end subroutine show
end module ar_mod

program test
    use ar_mod
    integer :: m(4,4)
    call show(m)
end program test
"#,
    );
}

// ── SELECT RANK construct (Fortran 2018) ──────────────────────

#[test]
fn select_rank_basic() {
    compile_ok(
        r#"
program test
    call handle([1, 2, 3])
contains
    subroutine handle(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(1)
            print *, 'rank-1 array, size =', size(x)
        rank default
            print *, 'other rank:', rank(x)
        end select
    end subroutine handle
end program test
"#,
    );
}

#[test]
fn select_rank_zero() {
    compile_ok(
        r#"
program test
    integer :: s = 99
    call inspect(s)
contains
    subroutine inspect(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(0)
            print *, 'scalar =', x
        rank(1)
            print *, 'vector'
        end select
    end subroutine inspect
end program test
"#,
    );
}

// ── IMPLICIT NONE (type, external) — F2018 extension ─────────

#[test]
fn implicit_none_type_external() {
    compile_ok(
        r#"
program test
    implicit none (type, external)
    integer :: x = 42
    print *, x
end program test
"#,
    );
}

// ── IS_CONTIGUOUS in various contexts (F2018 clarified) ───────

#[test]
fn is_contiguous_array() {
    compile_ok(
        r#"
program test
    real :: a(10)
    print *, is_contiguous(a)
end program test
"#,
    );
}

#[test]
fn is_contiguous_pointer_after_assign() {
    compile_ok(
        r#"
program test
    real, target :: a(10)
    real, pointer :: p(:)
    p => a
    print *, is_contiguous(p)
end program test
"#,
    );
}

#[test]
fn is_contiguous_non_unit_stride() {
    compile_ok(
        r#"
program test
    real, target :: a(10)
    real, pointer :: p(:)
    p => a(1:10:2)
    print *, is_contiguous(p)
end program test
"#,
    );
}

// ── Intrinsic module IEEE_ARITHMETIC extras (F2018) ──────────

#[test]
fn ieee_support_halting() {
    compile_ok(
        r#"
program test
    use ieee_arithmetic
    print *, ieee_support_halting(ieee_divide_by_zero)
end program test
"#,
    );
}

#[test]
fn ieee_support_denormal() {
    compile_ok(
        r#"
program test
    use ieee_arithmetic
    real :: x
    print *, ieee_support_denormal(x)
end program test
"#,
    );
}

#[test]
fn ieee_support_inf() {
    compile_ok(
        r#"
program test
    use ieee_arithmetic
    real :: x
    print *, ieee_support_inf(x)
end program test
"#,
    );
}

#[test]
fn ieee_support_nan() {
    compile_ok(
        r#"
program test
    use ieee_arithmetic
    real :: x
    print *, ieee_support_nan(x)
end program test
"#,
    );
}

// ── Coarray intrinsic functions (preview, covered in coarrays)

#[test]
fn image_index_intrinsic() {
    compile_ok(
        r#"
program test
    integer :: idx
    integer :: sub(2) = [1, 1]
    idx = image_index([2,2], sub)
    print *, idx
end program test
"#,
    );
}

// ── Error STOP with integer (F2018 allows integer) ────────────

#[test]
fn error_stop_integer() {
    compile_ok(
        r#"
program test
    logical :: ok = .true.
    if (.not. ok) error stop 1
    print *, 'fine'
end program test
"#,
    );
}

#[test]
fn error_stop_expression() {
    compile_ok(
        r#"
program test
    integer :: code = 0
    if (code /= 0) error stop code
    print *, 'ok'
end program test
"#,
    );
}

// ── STOP with integer expression (F2018 clarified) ───────────

#[test]
fn stop_zero() {
    compile_ok(
        r#"
program test
    print *, 'before'
    stop 0
end program test
"#,
    );
}

#[test]
fn stop_variable() {
    compile_ok(
        r#"
program test
    integer :: code = 0
    print *, 'ok'
    stop code
end program test
"#,
    );
}

// ── Intrinsic TYPEOF / RANK function (F2018) ─────────────────

#[test]
fn rank_intrinsic_scalar() {
    compile_ok(
        r#"
program test
    integer :: x = 5
    print *, rank(x)
end program test
"#,
    );
}

#[test]
fn rank_intrinsic_1d() {
    compile_ok(
        r#"
program test
    integer :: a(5)
    print *, rank(a)
end program test
"#,
    );
}

#[test]
fn rank_intrinsic_3d() {
    compile_ok(
        r#"
program test
    real :: m(2,3,4)
    print *, rank(m)
end program test
"#,
    );
}

// ── Implied-type CHARACTER (F2018 clarified) ──────────────────

#[test]
fn char_implied_type_concat() {
    compile_ok(
        r#"
program test
    character(len=*), parameter :: greeting = 'Hello, ' // 'World!'
    print *, greeting
end program test
"#,
    );
}

// ── Intrinsic CO_* collective subroutines (stub — single image)

#[test]
fn co_sum_single_image() {
    compile_ok(
        r#"
program test
    integer :: x = 42
    call co_sum(x)
    print *, x
end program test
"#,
    );
}

#[test]
fn co_max_single_image() {
    compile_ok(
        r#"
program test
    real :: x = 3.14
    call co_max(x)
    print *, x
end program test
"#,
    );
}

#[test]
fn co_min_single_image() {
    compile_ok(
        r#"
program test
    integer :: x = 7
    call co_min(x)
    print *, x
end program test
"#,
    );
}

#[test]
fn co_broadcast_single_image() {
    compile_ok(
        r#"
program test
    integer :: x = 99
    call co_broadcast(x, source_image=1)
    print *, x
end program test
"#,
    );
}

// ── Extended WRITE/PRINT (F2018 requirements) ─────────────────

#[test]
fn write_logical_format_extended() {
    compile_ok(
        r#"
program test
    logical :: flags(3) = [.true., .false., .true.]
    write(*, '(3L5)') flags
end program test
"#,
    );
}

#[test]
fn write_dt_format() {
    compile_ok(
        r#"
program test
    type :: Point
        real :: x, y
    end type Point
    type(Point) :: p
    p%x = 1.0; p%y = 2.0
    write(*, '(DT"point"(2))') p
end program test
"#,
    );
}

// ── EVENT and TEAM types (F2018 coarray synchronization) ──────

#[test]
fn event_type_declaration() {
    compile_ok(
        r#"
program test
    use iso_fortran_env
    type(event_type) :: ev[*]
    print *, 'ok'
end program test
"#,
    );
}

#[test]
fn team_type_declaration() {
    compile_ok(
        r#"
program test
    use iso_fortran_env
    type(team_type) :: my_team
    print *, 'ok'
end program test
"#,
    );
}

#[test]
fn form_team_basic() {
    compile_ok(
        r#"
program test
    use iso_fortran_env
    type(team_type) :: t
    call form_team(1, t)
    print *, 'ok'
end program test
"#,
    );
}

#[test]
fn change_team_construct() {
    compile_ok(
        r#"
program test
    use iso_fortran_env
    type(team_type) :: t
    call form_team(1, t)
    change team (t)
        print *, this_image()
    end team
end program test
"#,
    );
}

// ── GET_TEAM / TEAM_NUMBER / TEAM_IMAGE (F2018) ───────────────

#[test]
fn get_team_initial() {
    compile_ok(
        r#"
program test
    use iso_fortran_env
    type(team_type) :: t
    t = get_team(initial_team)
    print *, 'ok'
end program test
"#,
    );
}

#[test]
fn team_number_intrinsic() {
    compile_ok(
        r#"
program test
    use iso_fortran_env
    type(team_type) :: t
    t = get_team()
    print *, team_number(t)
end program test
"#,
    );
}

// ── STAT= and ERRMSG= on ALLOCATE (F2018 enhancement) ────────

#[test]
fn allocate_stat_errmsg() {
    compile_ok(
        r#"
program test
    integer, allocatable :: a(:)
    integer :: stat
    character(len=100) :: errmsg
    allocate(a(10), stat=stat, errmsg=errmsg)
    if (stat /= 0) then
        print *, trim(errmsg)
    else
        print *, size(a)
    end if
    deallocate(a)
end program test
"#,
    );
}

// ── CRITICAL construct (F2018 clarified semantics) ────────────

#[test]
fn critical_basic() {
    compile_ok(
        r#"
program test
    integer :: shared = 0
    critical
        shared = shared + 1
    end critical
    print *, shared
end program test
"#,
    );
}

// ── LOCK / UNLOCK (F2015+ via coarrays) ──────────────────────

#[test]
fn lock_unlock_basic() {
    compile_ok(
        r#"
program test
    use iso_fortran_env
    integer(lock_type) :: lk[*]
    lock(lk)
    print *, 'locked'
    unlock(lk)
    print *, 'unlocked'
end program test
"#,
    );
}

#[test]
fn fortran_2018_reduce_runtime_sum_and_mask() {
    let out = run_prints(
        r#"
program t
    integer :: a(6) = [1, 2, 3, 4, 5, 6]
    integer :: total
    logical :: mask(6) = [.true., .false., .true., .false., .true., .false.]
    total = reduce(a, operator(+))
    print *, total
    print *, reduce(a, operator(+), mask=mask)
end program t
"#,
    );
    assert_eq!(out, vec!["21", "9"]);
}

#[test]
fn fortran_2018_select_rank_runtime() {
    let out = run_prints(
        r#"
program t
    call handle([1, 2, 3])
    call inspect(42)
contains
    subroutine handle(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(1)
            print *, 1
        rank default
            print *, 0
        end select
    end subroutine handle

    subroutine inspect(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(0)
            print *, 0
        rank default
            print *, -1
        end select
    end subroutine inspect
end program t
"#,
    );
    assert_eq!(out, vec!["1", "0"]);
}

#[test]
fn fortran_2018_out_of_range_runtime() {
    let out = run_prints(
        r#"
program t
    integer(kind=4) :: small = 100
    print *, out_of_range(small, 0_2)
    print *, out_of_range(3.14, 0)
end program t
"#,
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn fortran_2018_is_contiguous_runtime() {
    let out = run_prints(
        r#"
program t
    real, target :: a(10)
    real, pointer :: full(:)
    real, pointer :: stride(:)
    full => a
    stride => a(1:10:2)
    print *, is_contiguous(full)
    print *, is_contiguous(stride)
end program t
"#,
    );
    assert_eq!(out, vec!["1", "0"]);
}
