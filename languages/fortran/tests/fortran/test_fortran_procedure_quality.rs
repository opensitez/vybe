use super::helpers::run_prints;

#[test]
fn procedure_quality_function_return() {
    let out = run_prints(
        r#"
program procedure_quality_function_return
    integer :: value

    value = square(7)
    print *, value

contains

integer function square(x)
    integer, intent(in) :: x
    square = x * x
end function square
end program procedure_quality_function_return
"#,
    );
    assert_eq!(out, vec!["49"]);
}

#[test]
fn procedure_quality_subroutine_inout() {
    let out = run_prints(
        r#"
program procedure_quality_subroutine_inout
    integer :: left
    integer :: right
    integer :: output
    left = 4
    right = 5
    call add_pair(left, right, output)
    print *, output

contains
    subroutine add_pair(a, b, c)
        integer, intent(in) :: a
        integer, intent(in) :: b
        integer, intent(out) :: c
        c = a + b
    end subroutine add_pair
end program procedure_quality_subroutine_inout
"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn procedure_quality_nested_function_visibility() {
    let out = run_prints(
        r#"
program procedure_quality_nested_function_visibility
    integer :: result
    call show_add(12, 8, result)
    print *, result

contains
    subroutine show_add(a, b, result)
        integer, intent(in) :: a, b
        integer, intent(out) :: result
        result = inc(a) + inc(b)
    contains
        integer function inc(v)
            integer, intent(in) :: v
            inc = v + 1
        end function inc
    end subroutine show_add
end program procedure_quality_nested_function_visibility
"#,
    );
    assert_eq!(out, vec!["22"]);
}

#[test]
fn procedure_quality_optional_argument_default() {
    let out = run_prints(
        r#"
program procedure_quality_optional_argument_default
    integer :: out_a
    integer :: out_b
    call multiply(3, out_a)
    call multiply(3, out_b, 4)
    print *, out_a
    print *, out_b

contains
    subroutine multiply(a, result, scale)
        integer, intent(in) :: a
        integer, intent(out) :: result
        integer, intent(in), optional :: scale
        if (present(scale)) then
            result = a * scale
        else
            result = a * 2
        end if
    end subroutine multiply
end program procedure_quality_optional_argument_default
"#,
    );
    assert_eq!(out, vec!["6", "12"]);
}

#[test]
fn procedure_quality_result_keyword() {
    let out = run_prints(
        r#"
program procedure_quality_result_keyword
    real :: ratio
    ratio = safe_ratio(10, 4)
    print *, ratio

contains
    real function safe_ratio(a, b) result(r)
        integer, intent(in) :: a
        integer, intent(in) :: b
        r = real(a) / real(b)
    end function safe_ratio
end program procedure_quality_result_keyword
"#,
    );
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn procedure_quality_interface_specificity() {
    let out = run_prints(
        r#"
program procedure_quality_interface_specificity
    integer :: int_out
    call call_increment(int_out)
    print *, int_out

contains
    interface
        function int_add_one(v) result(r)
            integer :: r
            integer, intent(in) :: v
        end function int_add_one
    end interface

    function int_add_one(v) result(r)
        integer, intent(in) :: v
        integer :: r
        r = v + 1
    end function int_add_one

    subroutine call_increment(result)
        integer, intent(out) :: result
        result = int_add_one(6)
    end subroutine call_increment
end program procedure_quality_interface_specificity
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn procedure_quality_array_argument_sum() {
    let out = run_prints(
        r#"
program procedure_quality_array_argument_sum
    integer, dimension(5) :: values
    integer :: total
    values = (/ 1, 2, 3, 4, 5 /)
    call array_sum(values, total)
    print *, total

contains
    subroutine array_sum(v, total)
        integer, intent(in) :: v(:)
        integer, intent(out) :: total
        integer :: i
        total = 0
        do i = 1, size(v)
            total = total + v(i)
        end do
    end subroutine array_sum
end program procedure_quality_array_argument_sum
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn procedure_quality_pure_function() {
    let out = run_prints(
        r#"
program procedure_quality_pure_function
    integer :: output
    output = negate_if_negative(-3)
    print *, output

contains
    integer function negate_if_negative(v)
        integer, intent(in) :: v
        if (v < 0) negate_if_negative = -v
        if (v >= 0) negate_if_negative = v
    end function negate_if_negative
end program procedure_quality_pure_function
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn procedure_quality_elemental_like_loop() {
    let out = run_prints(
        r#"
program procedure_quality_elemental_like_loop
    integer :: i
    integer :: out
    out = 1
    do i = 1, 5
        call set_scale(i, out)
    end do
    print *, out

contains
    subroutine set_scale(v, state)
        integer, intent(in) :: v
        integer, intent(inout) :: state
        state = state * v
    end subroutine set_scale
end program procedure_quality_elemental_like_loop
"#,
    );
    assert_eq!(out, vec!["120"]);
}

#[test]
fn procedure_quality_return_by_value_multiple() {
    let out = run_prints(
        r#"
program procedure_quality_return_by_value_multiple
    integer :: first
    integer :: second
    call minmax(8, 3, first, second)
    print *, first
    print *, second

contains
    subroutine minmax(a, b, lo, hi)
        integer, intent(in) :: a
        integer, intent(in) :: b
        integer, intent(out) :: lo
        integer, intent(out) :: hi
        if (a < b) then
            lo = a
            hi = b
        else
            lo = b
            hi = a
        end if
    end subroutine minmax
end program procedure_quality_return_by_value_multiple
"#,
    );
    assert_eq!(out, vec!["3", "8"]);
}

#[test]
fn procedure_quality_internal_state_update() {
    let out = run_prints(
        r#"
module counter_module
    integer, save :: counter = 0

    contains
    subroutine bump()
        counter = counter + 1
    end subroutine bump

    integer function value()
        value = counter
    end function value
end module counter_module

program procedure_quality_internal_state_update
    use counter_module
    call bump()
    call bump()
    print *, value()
end program procedure_quality_internal_state_update
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn procedure_quality_host_association() {
    let out = run_prints(
        r#"
program procedure_quality_host_association
    integer :: result
    integer :: base
    base = 4
    call caller(base, result)
    print *, result

contains
    subroutine caller(x, out)
        integer, intent(in) :: x
        integer, intent(out) :: out
        out = double(x)
    end subroutine caller

    integer function double(v)
        integer, intent(in) :: v
        double = v * 2
    end function double
end program procedure_quality_host_association
"#,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn procedure_quality_recursive_like_depth() {
    let out = run_prints(
        r#"
program procedure_quality_recursive_like_depth
    integer :: answer
    answer = countdown(4)
    print *, answer

contains
    integer function countdown(n)
        integer, intent(in) :: n
        if (n <= 0) then
            countdown = 0
        else
            countdown = n + countdown(n - 1)
        end if
    end function countdown
end program procedure_quality_recursive_like_depth
"#,
    );
    assert_eq!(out, vec!["10"]);
}
