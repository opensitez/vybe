use super::helpers::run_prints;

#[test]
fn subroutine_argument_default_values_optional_integer_simulated_default() {
    let out = run_prints(
        r#"
program subroutine_argument_default_values_optional_integer_simulated_default
    print *, apply_default(5)
    print *, apply_default(5, 2)
contains
    integer function apply_default(value, step)
        integer, intent(in) :: value
        integer, intent(in), optional :: step
        integer :: step_value
        if (present(step)) then
            step_value = step
        else
            step_value = 1
        end if
        apply_default = value + step_value
    end function apply_default
end program subroutine_argument_default_values_optional_integer_simulated_default
"#,
    );
    assert_eq!(out, vec!["6", "7"]);
}

#[test]
fn subroutine_argument_default_values_character_defaults() {
    let out = run_prints(
        r#"
program subroutine_argument_default_values_character_defaults
    character(len=16) :: a
    character(len=16) :: b
    call build_label(a)
    call build_label(b, 'x')
    print *, trim(a)
    print *, trim(b)
contains
    subroutine build_label(out, suffix)
        character(len=*), intent(out) :: out
        character(len=*), intent(in), optional :: suffix
        if (present(suffix)) then
            out = trim('root') // trim(suffix)
        else
            out = 'root'
        end if
    end subroutine build_label
end program subroutine_argument_default_values_character_defaults
"#,
    );
    assert_eq!(out, vec!["root", "rootx"]);
}

#[test]
fn subroutine_argument_default_values_logical_defaults() {
    let out = run_prints(
        r#"
program subroutine_argument_default_values_logical_defaults
    logical :: a
    logical :: b
    call toggle(a)
    call toggle(b, .false.)
    print *, a
    print *, b
contains
    subroutine toggle(value, active)
        logical, intent(out) :: value
        logical, intent(in), optional :: active
        if (present(active)) then
            value = active
        else
            value = .true.
        end if
    end subroutine toggle
end program subroutine_argument_default_values_logical_defaults
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn subroutine_argument_default_values_real_defaults() {
    let out = run_prints(
        r#"
program subroutine_argument_default_values_real_defaults
    print *, ratio(10.0)
    print *, ratio(10.0, 4.0)
contains
    real function ratio(value, divisor)
        real, intent(in) :: value
        real, intent(in), optional :: divisor
        real :: d
        if (present(divisor)) then
            d = divisor
        else
            d = 2.0
        end if
        ratio = value / d
    end function ratio
end program subroutine_argument_default_values_real_defaults
"#,
    );
    assert_eq!(out, vec!["5", "2"]);
}

#[test]
fn subroutine_argument_default_values_array_defaults_with_optional_shape() {
    let out = run_prints(
        r#"
program subroutine_argument_default_values_array_defaults_with_optional_shape
    integer :: a(3)
    call fill(a)
    print *, sum(a)
contains
    subroutine fill(out, scale)
        integer, intent(out) :: out(:)
        integer, intent(in), optional :: scale
        integer :: i
        integer :: factor
        if (present(scale)) then
            factor = scale
        else
            factor = 1
        end if
        do i = 1, size(out)
            out(i) = i * factor
        end do
    end subroutine fill
end program subroutine_argument_default_values_array_defaults_with_optional_shape
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn subroutine_argument_default_values_type_defaults_via_optional() {
    let out = run_prints(
        r#"
program subroutine_argument_default_values_type_defaults_via_optional
    type item
        integer :: x
    end type item
    type(item) :: base
    print *, set_item(base, 4)%x
contains
    function set_item(base, x) result(out)
        type(item), intent(in) :: base
        integer, intent(in), optional :: x
        type(item) :: out
        if (present(x)) then
            out%x = base%x + x
        else
            out%x = base%x
        end if
    end function set_item
end program subroutine_argument_default_values_type_defaults_via_optional
"#,
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn subroutine_argument_default_values_keyword_defaults() {
    let out = run_prints(
        r#"
program subroutine_argument_default_values_keyword_defaults
    print *, combine(a=1)
    print *, combine(a=1, b=2)
contains
    integer function combine(a, b, c)
        integer, intent(in) :: a
        integer, intent(in), optional :: b, c
        combine = a
        if (present(b)) combine = combine + b
        if (present(c)) combine = combine + c
    end function combine
end program subroutine_argument_default_values_keyword_defaults
"#,
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn subroutine_argument_default_values_optional_out_with_default_mode() {
    let out = run_prints(
        r#"
program subroutine_argument_default_values_optional_out_with_default_mode
    integer :: out_val
    call maybe_set(out_val)
    print *, out_val
contains
    subroutine maybe_set(result, value)
        integer, intent(out) :: result
        integer, intent(in), optional :: value
        if (present(value)) then
            result = value
        else
            result = 99
        end if
    end subroutine maybe_set
end program subroutine_argument_default_values_optional_out_with_default_mode
"#,
    );
    assert_eq!(out, vec!["99"]);
}
