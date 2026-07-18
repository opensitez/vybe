use super::helpers::run_prints;

fn assert_single_i64_print(src: String, expected: i64) {
    assert_eq!(run_prints(&src), vec![expected.to_string()]);
}

fn assert_double_i64_print(src: String, expected: (i64, i64)) {
    assert_eq!(
        run_prints(&src),
        vec![expected.0.to_string(), expected.1.to_string()],
    );
}

#[test]
fn fortran_matrix_arithmetic_expressions() {
    for n in 1..=100 {
        let expected = (n as i64) * (n as i64) - (n as i64) + 3;
        let src = format!(
            r#"
program fortran_matrix_arithmetic_expressions
    integer :: n
    n = {n}
    print *, n * n - n + 3
end program fortran_matrix_arithmetic_expressions
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_branching_by_mod() {
    for n in 1..=100 {
        let expected = if n % 3 == 0 { (n / 3) as i64 } else { (n - 1) as i64 };
        let src = format!(
            r#"
program fortran_matrix_branching_by_mod
    integer :: n
    n = {n}
    if (mod(n, 3) == 0) then
        print *, n / 3
    else
        print *, n - 1
    end if
end program fortran_matrix_branching_by_mod
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_triangle_sums() {
    for n in 1..=100 {
        let expected = (n as i64 * (n as i64 + 1)) / 2;
        let src = format!(
            r#"
program fortran_matrix_triangle_sums
    integer :: n, i
    integer :: total
    total = 0
    do i = 1, {n}
        total = total + i
    end do
    print *, total
end program fortran_matrix_triangle_sums
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_array_section_sizes() {
    for n in 1..=100 {
        let start = (n % 10) + 1;
        let expected = ((10 - start) / 2 + 1) as i64;
        let src = format!(
            r#"
program fortran_matrix_array_section_sizes
    integer :: a(1:10)
    integer :: n
    n = {n}
    a = (/1,2,3,4,5,6,7,8,9,10/)
    print *, size(a(mod(n-1,10)+1:10:2))
end program fortran_matrix_array_section_sizes
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_where_mask_updates() {
    for n in 1..=100 {
        let expected = 55 + (n.min(10) as i64);
        let src = format!(
            r#"
program fortran_matrix_where_mask_updates
    integer :: values(1:10)
    integer :: n
    n = {n}
    values = (/1,2,3,4,5,6,7,8,9,10/)
    where (values <= n)
        values = values + 1
    end where
    print *, sum(values)
end program fortran_matrix_where_mask_updates
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_nested_boolean_categorization() {
    for n in 1..=100 {
        let mut expected = 3;
        if n < 20 {
            expected = 1;
        } else if n < 60 {
            expected = 2;
        }
        if n % 2 == 0 {
            expected += 10;
        }
        let src = format!(
            r#"
program fortran_matrix_nested_boolean_categorization
    integer :: n
    integer :: class
    n = {n}
    if (n < 20) then
        class = 1
    else if (n < 60) then
        class = 2
    else
        class = 3
    end if
    if (mod(n,2) == 0) then
        class = class + 10
    end if
    print *, class
end program fortran_matrix_nested_boolean_categorization
"#,
        );
        assert_single_i64_print(src, expected as i64);
    }
}

#[test]
fn fortran_matrix_select_case_ranges() {
    for n in 1..=100 {
        let expected = if n <= 30 {
            1
        } else if n <= 60 {
            2
        } else {
            3
        } as i64;
        let src = format!(
            r#"
program fortran_matrix_select_case_ranges
    integer :: n
    integer :: bucket
    n = {n}
    select case (n)
    case (1:30)
        bucket = 1
    case (31:60)
        bucket = 2
    case default
        bucket = 3
    end select
    print *, bucket
end program fortran_matrix_select_case_ranges
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_character_code_sequences() {
    for n in 1..=100 {
        let expected = (65 + ((n - 1) % 26)) as i64;
        let src = format!(
            r#"
program fortran_matrix_character_code_sequences
    integer :: code
    code = iachar('A') + mod({n} - 1, 26)
    print *, code
end program fortran_matrix_character_code_sequences
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_derived_type_addition() {
    for n in 1..=100 {
        let expected = 3 * n as i64;
        let src = format!(
            r#"
program fortran_matrix_derived_type_addition
    type pair
        integer :: left
        integer :: right
    end type pair
    type(pair) :: p
    p%left = {n}
    p%right = {n} * 2
    print *, p%left + p%right
end program fortran_matrix_derived_type_addition
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_int_truncation_paths() {
    for n in 1..=100 {
        let expected = ((n as f64) / 2.0) as i64;
        let src = format!(
            r#"
program fortran_matrix_int_truncation_paths
    integer :: n
    real :: as_real
    n = {n}
    as_real = real(n) / 2.0
    print *, int(as_real)
end program fortran_matrix_int_truncation_paths
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_module_factor_scaling() {
    for n in 1..=100 {
        let expected = 3 * n as i64 + 1;
        let src = format!(
            r#"
module fortran_matrix_scale_module
    implicit none
    integer, parameter :: module_factor = 3
contains
    integer function scaled(v)
        integer, intent(in) :: v
        scaled = v * module_factor + 1
    end function scaled
end module fortran_matrix_scale_module

program fortran_matrix_module_factor_scaling
    use fortran_matrix_scale_module
    print *, scaled({n})
end program fortran_matrix_module_factor_scaling
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_optional_argument_results() {
    for n in 1..=100 {
        let expected = if n % 2 == 0 {
            (n + 1) as i64
        } else {
            (n + 3) as i64
        };
        let src = format!(
            r#"
program fortran_matrix_optional_argument_results
    integer :: v, out
    v = {n}
    if (mod(v, 2) == 0) then
        call emit(v, .true., out)
    else
        call emit(v, out=out)
    end if
    print *, out
contains
    subroutine emit(v, use_double, out)
        integer, intent(in) :: v
        logical, optional, intent(in) :: use_double
        integer, intent(out) :: out
        if (present(use_double)) then
            out = v + 1
        else
            out = v + 3
        end if
    end subroutine emit
end program fortran_matrix_optional_argument_results
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_pointer_aliasing_updates() {
    for n in 1..=100 {
        let expected = 2 * n as i64 + 2;
        let src = format!(
            r#"
program fortran_matrix_pointer_aliasing_updates
    integer, target :: container(2)
    integer, pointer :: view(:)
    container = (/ {n}, {n} + 1 /)
    view => container
    view(1) = view(1) + 1
    print *, container(1) + container(2)
end program fortran_matrix_pointer_aliasing_updates
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_procedure_pointer_dispatch() {
    for n in 1..=100 {
        let expected = if n % 2 == 0 { 2 * n } else { 3 * n } as i64;
        let src = format!(
            r#"
program fortran_matrix_procedure_pointer_dispatch
    abstract interface
        integer function transform(value)
            integer, intent(in) :: value
        end function transform
    end interface
    procedure(transform), pointer :: fp

    if (mod({n}, 2) == 0) then
        fp => double
    else
        fp => triple
    end if
    print *, fp({n})

contains
    integer function double(value)
        integer, intent(in) :: value
        double = value * 2
    end function double

    integer function triple(value)
        integer, intent(in) :: value
        triple = value * 3
    end function triple
end program fortran_matrix_procedure_pointer_dispatch
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_dynamic_selection_type() {
    for n in 1..=100 {
        let expected = if n % 2 == 0 { (10 + n) } else { (20 + n) } as i64;
        let src = format!(
            r#"
program fortran_matrix_dynamic_selection_type
    class(*), allocatable :: x
    integer :: n
    n = {n}

    if (mod(n, 2) == 0) then
        allocate(integer :: x)
    else
        allocate(real :: x)
    end if

    select type (x)
    type is (integer)
        x = n
        print *, n + 10
    type is (real)
        x = real(n)
        print *, n + 20
    class default
        print *, 0
    end select
end program fortran_matrix_dynamic_selection_type
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_block_data_scoping() {
    for n in 1..=100 {
        let expected = (2 * n) as i64;
        let src = format!(
            r#"
program fortran_matrix_block_data_scoping
    integer :: outer
    outer = {n} * 2
    if (mod(outer, 2) == 0) then
        block
            integer :: inner
            inner = outer + 2
        end block
    end if
    print *, outer
end program fortran_matrix_block_data_scoping
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_array_extremes_by_stride() {
    for n in 1..=100 {
        let expected = (n as i64 * 3 + 3);
        let src = format!(
            r#"
program fortran_matrix_array_extremes_by_stride
    integer :: values(1:3)
    integer :: n
    n = {n}
    values = (/n, n+1, n+2/)
    print *, sum(values)
end program fortran_matrix_array_extremes_by_stride
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_recursive_termination_small_factorial() {
    for n in 1..=100 {
        let reduced = (n - 1) % 8 + 1;
        let mut expected = 1i64;
        for value in 2..=reduced as i64 {
            expected *= value;
        }
        let src = format!(
            r#"
program fortran_matrix_recursive_termination_small_factorial
    integer :: n
    integer :: result
    n = mod({n}-1, 8) + 1
    result = fact(n)
    print *, result

contains
    recursive integer function fact(value)
        integer, intent(in) :: value
        if (value <= 1) then
            fact = 1
        else
            fact = value * fact(value - 1)
        end if
    end function fact
end program fortran_matrix_recursive_termination_small_factorial
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_string_length_probe() {
    for n in 1..=100 {
        let expected = (3 + ((n - 1) % 5)) as i64;
        let src = format!(
            r#"
program fortran_matrix_string_length_probe
    character(len=10) :: s
    integer :: n
    n = {n}
    select case (mod(n-1, 5))
    case (0)
        s = "abc"
    case (1)
        s = "abcd"
    case (2)
        s = "abcde"
    case (3)
        s = "abcdef"
    case (4)
        s = "abcdefg"
    end select
    print *, len_trim(s)
end program fortran_matrix_string_length_probe
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_logical_gate_table() {
    for n in 1..=100 {
        let expected = match n % 4 {
            1 => 0,
            2 => 2,
            3 => 4,
            _ => 6,
        };
        let src = format!(
            r#"
program fortran_matrix_logical_gate_table
    integer :: n
    integer :: out
    n = {n}
    if (mod(n,2) == 0) then
        if (mod(n,4) == 0) then
            out = 6
        else
            out = 2
        end if
    else
        if (mod(n,4) == 1) then
            out = 0
        else
            out = 4
        end if
    end if
    print *, out
end program fortran_matrix_logical_gate_table
"#,
        );
        assert_single_i64_print(src, expected);
    }
}

#[test]
fn fortran_matrix_minmax_pairing() {
    for n in 1..=100 {
        let expected_min = n as i64;
        let expected_max = if n < 5 { (n + 5) as i64 } else { (3 * n) as i64 };
        let src = format!(
            r#"
program fortran_matrix_minmax_pairing
    integer :: n
    integer :: values(3)
    n = {n}
    values = (/n, n*2, n+5/)
    print *, minval(values)
    print *, maxval(values)
end program fortran_matrix_minmax_pairing
"#,
        );
        assert_double_i64_print(src, (expected_min, expected_max));
    }
}
