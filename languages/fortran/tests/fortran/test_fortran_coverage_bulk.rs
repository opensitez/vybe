use super::helpers::run_prints;

#[test]
fn fortran_bulk_array_section_mass() {
    let out = run_prints(
        r#"
program fortran_bulk_array_section_mass
    integer :: values(1:200), i
    values = (/ (i, i = 1, 200) /)
    do i = 1, 100
        print *, sum(values(i:i+10))
    end do
end program fortran_bulk_array_section_mass
"#,
    );

    assert_eq!(
        out,
        (1..=100)
            .map(|i| (11 * (2 * i + 9) / 2).to_string())
            .collect::<Vec<String>>()
    );
}

#[test]
fn fortran_bulk_array_shape_casts() {
    let out = run_prints(
        r#"
program fortran_bulk_array_shape_casts
    integer :: source(10)
    integer :: matrix(2,5)
    integer :: i
    source = (/ (i, i = 1, 10) /)
    matrix = reshape(source, shape(matrix))
    do i = 1, 10
        print *, matrix(mod(i-1,2)+1, (i-1)/2 + 1)
    end do
end program fortran_bulk_array_shape_casts
"#,
    );

    assert_eq!(
        out,
        (1..=10).map(|i| i.to_string()).collect::<Vec<String>>()
    );
}

#[test]
fn fortran_bulk_string_intrinsics_chain() {
    let out = run_prints(
        r#"
program fortran_bulk_string_intrinsics_chain
    character(len=20) :: s
    character(len=20) :: t
    integer :: i
    do i = 1, 100
        s = ' item-' // trim(adjustl(transfer(i, '       ')))
        t = trim(adjustl(s)) // '-' // trim(transfer(i, ''))
        print *, len_trim(t)
    end do
end program fortran_bulk_string_intrinsics_chain
"#,
    );

    assert_eq!(
        out,
        (1..=100)
            .map(|i| {
                let t = format!("item-{}-{}", i, i);
                (t.len()).to_string()
            })
            .collect::<Vec<String>>(),
    );
}

#[test]
fn fortran_bulk_select_case_mixed() {
    let out = run_prints(
        r#"
program fortran_bulk_select_case_mixed
    integer :: i
    do i = 1, 100
        select case (i)
        case (1:33)
            print *, 'low'
        case (34:66)
            print *, 'mid'
        case default
            print *, 'hi'
        end select
    end do
end program fortran_bulk_select_case_mixed
"#,
    );

    let expected: Vec<String> = (1..=100)
        .map(|i| {
            if i <= 33 {
                "low".to_string()
            } else if i <= 66 {
                "mid".to_string()
            } else {
                "hi".to_string()
            }
        })
        .collect::<Vec<String>>();
    assert_eq!(out, expected);
}

#[test]
fn fortran_bulk_select_case_character_ranges() {
    let out = run_prints(
        r#"
program fortran_bulk_select_case_character_ranges
    integer :: i
    character(len=1) :: c
    do i = 1, 100
        c = achar(iachar('a') + mod(i-1, 26))
        select case (c)
        case ('a':'m')
            print *, 'first'
        case ('n':'z')
            print *, 'second'
        case default
            print *, 'wrap'
        end select
    end do
end program fortran_bulk_select_case_character_ranges
"#,
    );

    let expected: Vec<String> = (1..=100)
        .map(|i| {
            if ((i - 1) % 26 + 1) <= 13 {
                "first"
            } else {
                "second"
            }
        })
        .map(|x| x.to_string())
        .collect::<Vec<String>>();
    assert_eq!(out, expected);
}

#[test]
fn fortran_bulk_subroutine_optionals_flow() {
    let out = run_prints(
        r#"
program fortran_bulk_subroutine_optionals_flow
    integer :: i
    do i = 1, 100
        print *, opt_sum(i)
    end do
contains
    integer function opt_sum(v)
        integer, intent(in) :: v
        integer, optional :: bump
        if (present(bump)) then
            opt_sum = v + bump
        else
            opt_sum = v
        end if
    end function opt_sum
end program fortran_bulk_subroutine_optionals_flow
"#,
    );

    assert_eq!(
        out,
        (1..=100).map(|i| i.to_string()).collect::<Vec<String>>()
    );
}

#[test]
fn fortran_bulk_pointer_reassociation() {
    let out = run_prints(
        r#"
program fortran_bulk_pointer_reassociation
    integer, target :: base(100)
    integer, pointer :: p, q
    integer :: i
    do i = 1, 100
        base(i) = i
    end do
    do i = 1, 100
        p => base(i:i)
        q => p
        print *, q(1)
    end do
end program fortran_bulk_pointer_reassociation
"#,
    );

    assert_eq!(
        out,
        (1..=100).map(|i| i.to_string()).collect::<Vec<String>>()
    );
}

#[test]
fn fortran_bulk_procedure_pointer_dispatch() {
    let out = run_prints(
        r#"
program fortran_bulk_procedure_pointer_dispatch
    integer :: i
    abstract interface
        integer function calc(a)
            integer, intent(in) :: a
        end function calc
    end interface
    procedure(calc), pointer :: f

    do i = 1, 50
        if (mod(i,2) == 0) then
            f => double_it
        else
            f => triple_it
        end if
        print *, f(i)
    end do
contains
    integer function double_it(v)
        integer, intent(in) :: v
        double_it = 2 * v
    end function double_it
    integer function triple_it(v)
        integer, intent(in) :: v
        triple_it = 3 * v
    end function triple_it
end program fortran_bulk_procedure_pointer_dispatch
"#,
    );

    assert_eq!(
        out,
        (1..=50)
            .map(|i| if i % 2 == 0 {
                (2 * i).to_string()
            } else {
                (3 * i).to_string()
            })
            .collect::<Vec<String>>(),
    );
}

#[test]
fn fortran_bulk_recursive_factorials() {
    let out = run_prints(
        r#"
program fortran_bulk_recursive_factorials
    integer :: i
    do i = 1, 40
        print *, fact(min(i, 10))
    end do
contains
    recursive integer function fact(n)
        integer, intent(in) :: n
        if (n <= 1) then
            fact = 1
        else
            fact = n * fact(n - 1)
        end if
    end function fact
end program fortran_bulk_recursive_factorials
"#,
    );

    let expected: Vec<String> = (1..=40)
        .map(|i| {
            let n = if i <= 10 { i } else { 10 };
            let mut r = 1;
            for v in 2..=n {
                r *= v;
            }
            r.to_string()
        })
        .collect::<Vec<String>>();
    assert_eq!(out, expected);
}

#[test]
fn fortran_bulk_logical_parentheses() {
    let out = run_prints(
        r#"
program fortran_bulk_logical_parentheses
    integer :: i
    logical :: a, b
    do i = 1, 100
        a = (mod(i,3) == 0)
        b = (mod(i,5) == 0)
        if ((a .and. .not. b) .or. (.not. a .and. b)) then
            print *, 'xor'
        else
            print *, 'eq'
        end if
    end do
end program fortran_bulk_logical_parentheses
"#,
    );

    let expected: Vec<String> = (1..=100)
        .map(|i| {
            let a = i % 3 == 0;
            let b = i % 5 == 0;
            if (a && !b) || (!a && b) { "xor" } else { "eq" }
        })
        .map(|s| s.to_string())
        .collect::<Vec<String>>();
    assert_eq!(out, expected);
}

#[test]
fn fortran_bulk_implicit_none_recovery() {
    let out = run_prints(
        r#"
program fortran_bulk_implicit_none_recovery
    implicit none
    integer :: i
    do i = 1, 100
        print *, i + 2
    end do
end program fortran_bulk_implicit_none_recovery
"#,
    );

    assert_eq!(
        out,
        (3..=102).map(|v| v.to_string()).collect::<Vec<String>>()
    );
}

#[test]
fn fortran_bulk_namespace_and_modules() {
    let out = run_prints(
        r#"
module bulk_ns
    integer, parameter :: base = 3
contains
    integer function scale(v)
        integer, intent(in) :: v
        scale = base * v
    end function scale
end module bulk_ns

program fortran_bulk_namespace_and_modules
    use bulk_ns
    integer :: i
    do i = 1, 50
        print *, scale(i)
    end do
end program fortran_bulk_namespace_and_modules
"#,
    );

    assert_eq!(
        out,
        (1..=50)
            .map(|v| (3 * v).to_string())
            .collect::<Vec<String>>(),
    );
}

#[test]
fn fortran_bulk_array_bounds_edges() {
    let out = run_prints(
        r#"
program fortran_bulk_array_bounds_edges
    integer :: a(-2:2)
    integer :: i
    do i = -2, 2
        a(i) = i + 10
    end do
    do i = -2, 2
        print *, a(i)
    end do
end program fortran_bulk_array_bounds_edges
"#,
    );

    assert_eq!(
        out,
        (8..=12).map(|v| v.to_string()).collect::<Vec<String>>()
    );
}

#[test]
fn fortran_bulk_associative_pointer_lifetimes() {
    let out = run_prints(
        r#"
program fortran_bulk_associative_pointer_lifetimes
    integer, allocatable, save :: box(:)
    integer, pointer :: p(:)
    integer :: i

    allocate(box(1:20))
    box = (/ (i, i = 1, 20) /)
    p => box
    do i = 1, 20
        print *, p(i)
    end do
end program fortran_bulk_associative_pointer_lifetimes
"#,
    );

    assert_eq!(
        out,
        (1..=20).map(|v| v.to_string()).collect::<Vec<String>>()
    );
}

#[test]
fn fortran_bulk_derived_type_defaults() {
    let out = run_prints(
        r#"
program fortran_bulk_derived_type_defaults
    type config
        integer :: a = 2
        integer :: b = 3
        integer :: c
    end type config
    type(config) :: x
    integer :: i
    x%c = x%a + x%b
    do i = 1, 100
        print *, x%c + i
    end do
end program fortran_bulk_derived_type_defaults
"#,
    );

    assert_eq!(
        out,
        (3..=102)
            .map(|v| (v + 2).to_string())
            .collect::<Vec<String>>()
    );
}

#[test]
fn fortran_bulk_do_construct_progress() {
    let out = run_prints(
        r#"
program fortran_bulk_do_construct_progress
    integer :: i, sum
    sum = 0
    do i = 1, 100, 2
        sum = sum + i
        print *, sum
    end do
end program fortran_bulk_do_construct_progress
"#,
    );

    let mut acc = 0;
    let expected = (1..=100)
        .step_by(2)
        .map(|v| {
            acc += v;
            acc
        })
        .map(|v| v.to_string())
        .collect::<Vec<String>>();

    assert_eq!(out, expected);
}

#[test]
fn fortran_bulk_select_type_class() {
    let out = run_prints(
        r#"
program fortran_bulk_select_type_class
    class(*), allocatable :: x
    integer :: i

    do i = 1, 100
        if (mod(i, 2) == 0) then
            allocate(real :: x)
        else
            allocate(integer :: x)
        end if
        x = i
        select type (x)
        type is (integer)
            print *, 1
        type is (real)
            print *, 2
        class default
            print *, 3
        end select
    end do
end program fortran_bulk_select_type_class
"#,
    );

    let expected: Vec<String> = (1..=100)
        .map(|i| if i % 2 == 0 { "2" } else { "1" }.to_string())
        .collect::<Vec<String>>();
    assert_eq!(out, expected);
}

#[test]
fn fortran_bulk_optional_keywords() {
    let out = run_prints(
        r#"
program fortran_bulk_optional_keywords
    integer :: i
    do i = 1, 100
        call emit(i)
    end do
contains
    subroutine emit(v, skip)
        integer, intent(in) :: v
        logical, optional, intent(in) :: skip
        if (present(skip) .and. skip) then
            print *, 'S'
        else
            print *, v
        end if
    end subroutine emit
end program fortran_bulk_optional_keywords
"#,
    );

    assert_eq!(
        out,
        (1..=100).map(|i| i.to_string()).collect::<Vec<String>>()
    );
}

#[test]
fn fortran_bulk_name_resolution_nested_blocks() {
    let out = run_prints(
        r#"
program fortran_bulk_name_resolution_nested_blocks
    integer :: outer
    outer = 10
    block
        integer :: outer
        outer = 20
        block
            integer :: inner
            inner = outer + 1
            print *, inner
        end block
        print *, outer
    end block
    print *, outer
end program fortran_bulk_name_resolution_nested_blocks
"#,
    );

    assert_eq!(out, vec!["21", "20", "10"]);
}

#[test]
fn fortran_bulk_file_unit_style() {
    let out = run_prints(
        r#"
program fortran_bulk_file_unit_style
    integer :: u
    integer :: i
    open(newunit=u, status='scratch', action='readwrite')
    do i = 1, 20
        write(u, '(I0)') i
    end do
    rewind(u)
    do i = 1, 20
        print *, i
    end do
    close(u)
end program fortran_bulk_file_unit_style
"#,
    );

    assert_eq!(
        out,
        (1..=20).map(|i| i.to_string()).collect::<Vec<String>>()
    );
}

#[test]
fn fortran_bulk_stop_error_paths() {
    let out = run_prints(
        r#"
program fortran_bulk_stop_error_paths
    integer :: i
    do i = 1, 100
        if (mod(i, 25) == 0) then
            print *, 0
        else
            print *, 1
        end if
    end do
end program fortran_bulk_stop_error_paths
"#,
    );

    assert_eq!(
        out,
        (1..=100)
            .map(|i| {
                if i % 25 == 0 {
                    "0".to_string()
                } else {
                    "1".to_string()
                }
            })
            .collect::<Vec<String>>(),
    );
}

#[test]
fn fortran_bulk_static_init_chain() {
    let out = run_prints(
        r#"
program fortran_bulk_static_init_chain
    integer, save :: a = 1
    integer, save :: b = a + 1
    integer, save :: c = b + 1
    integer :: i
    do i = 1, 100
        print *, a + b + c + i
    end do
end program fortran_bulk_static_init_chain
"#,
    );

    assert_eq!(
        out,
        (1..=100)
            .map(|i| (i + 6).to_string())
            .collect::<Vec<String>>()
    );
}

#[test]
fn fortran_bulk_floating_point_rounding_like() {
    let out = run_prints(
        r#"
program fortran_bulk_floating_point_rounding_like
    integer :: i
    real :: x
    do i = 1, 100
        x = 1.0 / real(i)
        print *, nint(x*10.0)
    end do
end program fortran_bulk_floating_point_rounding_like
"#,
    );

    let expected: Vec<String> = (1..=100)
        .map(|i| (10.0 / i as f64).round() as i32)
        .map(|v| v.to_string())
        .collect::<Vec<String>>();
    assert_eq!(out, expected);
}

#[test]
fn fortran_bulk_associate_array_aliasing() {
    let out = run_prints(
        r#"
program fortran_bulk_associate_array_aliasing
    integer :: values(4)
    integer :: i
    values = (/ 11, 22, 33, 44 /)

    associate (middle => values(2:3))
        middle = middle * 2
    end associate

    do i = 1, 4
        print *, values(i)
    end do
end program fortran_bulk_associate_array_aliasing
"#,
    );

    assert_eq!(
        out,
        vec![
            "11".to_string(),
            "44".to_string(),
            "66".to_string(),
            "44".to_string()
        ]
    );
}

#[test]
fn fortran_bulk_forall_masked_transform() {
    let out = run_prints(
        r#"
program fortran_bulk_forall_masked_transform
    integer :: values(10)
    integer :: i
    values = (/ (i, i = 1, 10) /)

    forall (i = 1:10, mod(i, 2) == 0)
        values(i) = values(i) + 100
    end forall

    do i = 1, 10
        print *, values(i)
    end do
end program fortran_bulk_forall_masked_transform
"#,
    );

    let expected = (1..=10)
        .map(|i| {
            if i % 2 == 0 {
                (i + 100).to_string()
            } else {
                i.to_string()
            }
        })
        .collect::<Vec<String>>();
    assert_eq!(out, expected);
}

#[test]
fn fortran_bulk_where_elsewhere_mix() {
    let out = run_prints(
        r#"
program fortran_bulk_where_elsewhere_mix
    integer :: source(1:8)
    integer :: transformed(1:8)
    integer :: i
    do i = 1, 8
        source(i) = i
    end do

    where (mod(source, 3) == 0)
        transformed = source * 10
    elsewhere
        transformed = source + 1
    end where

    do i = 1, 8
        print *, transformed(i)
    end do
end program fortran_bulk_where_elsewhere_mix
"#,
    );

    let expected: Vec<String> = (1..=8)
        .map(|i| if i % 3 == 0 { i * 10 } else { i + 1 })
        .map(|v| v.to_string())
        .collect::<Vec<String>>();
    assert_eq!(out, expected);
}

#[test]
fn fortran_bulk_nested_associate_array_aliasing() {
    let out = run_prints(
        r#"
program fortran_bulk_nested_associate_array_aliasing
    integer :: buffer(4)
    integer :: base

    base = 5

    associate (whole => buffer)
        whole = (/ 1, 2, 3, 4 /)
        associate (edge => whole(2:3))
            edge = edge + base
        end associate
        print *, whole(1), whole(2), whole(3), whole(4)
    end associate
end program fortran_bulk_nested_associate_array_aliasing
"#,
    );

    assert_eq!(out, vec!["1", "7", "8", "4"]);
}

#[test]
fn fortran_bulk_forall_multi_index_matrix() {
    let out = run_prints(
        r#"
program fortran_bulk_forall_multi_index_matrix
    integer :: m(4,3)
    integer :: i, j

    do i = 1, 4
        do j = 1, 3
            m(i, j) = 10 * i + j
        end do
    end do

    forall (i = 1:4, j = 1:3, i == j)
        m(i, j) = m(i, j) + 100
    end forall

    do i = 1, 4
        do j = 1, 3
            print *, m(i, j)
        end do
    end do
end program fortran_bulk_forall_multi_index_matrix
"#,
    );

    let expected = vec![
        "111", "12", "13", //
        "21", "122", "23", //
        "31", "32", "133", //
        "41", "42", "43",
    ]
    .iter()
    .map(|s| s.to_string())
    .collect::<Vec<String>>();
    assert_eq!(out, expected);
}

#[test]
fn fortran_bulk_where_logical_mask_array() {
    let out = run_prints(
        r#"
program fortran_bulk_where_logical_mask_array
    integer :: a(1:8)
    integer :: b(1:8)
    logical :: is_even(1:8)
    integer :: i

    do i = 1, 8
        a(i) = i
        is_even(i) = mod(i, 2) == 0
    end do

    where (is_even)
        b = a * 3
    elsewhere
        b = a + 10
    end where

    do i = 1, 8
        print *, b(i)
    end do
end program fortran_bulk_where_logical_mask_array
"#,
    );

    let expected = (1..=8)
        .map(|i| if i % 2 == 0 { (i * 3) } else { (i + 10) })
        .map(|v| v.to_string())
        .collect::<Vec<String>>();
    assert_eq!(out, expected);
}
