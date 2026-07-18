use super::helpers::run_prints;

#[test]
fn procedure_associated_pointer_calls_basic_subroutine_indirection() {
    let out = run_prints(
        r#"
program procedure_associated_pointer_calls_basic_subroutine_indirection
    abstract interface
        subroutine op(a, b, c)
            integer, intent(in) :: a, b
            integer, intent(out) :: c
        end subroutine op
    end interface

    procedure(op), pointer :: p
    integer :: result

    p => add_two
    call p(3, 4, result)
    print *, result
contains
    subroutine add_two(a, b, c)
        integer, intent(in) :: a, b
        integer, intent(out) :: c
        c = a + b
    end subroutine add_two
end program procedure_associated_pointer_calls_basic_subroutine_indirection
"#,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn procedure_associated_pointer_calls_function_indirection() {
    let out = run_prints(
        r#"
program procedure_associated_pointer_calls_function_indirection
    abstract interface
        integer function f(a, b)
            integer, intent(in) :: a, b
        end function f
    end interface

    procedure(f), pointer :: p
    integer :: result

    p => add_mul
    result = p(2, 3)
    print *, result
contains
    integer function add_mul(a, b)
        integer, intent(in) :: a, b
        add_mul = a * b + a
    end function add_mul
end program procedure_associated_pointer_calls_function_indirection
"#,
    );
    assert_eq!(out, vec!["8"]);
}

#[test]
fn procedure_associated_pointer_calls_pointer_reassignment() {
    let out = run_prints(
        r#"
program procedure_associated_pointer_calls_pointer_reassignment
    abstract interface
        integer function f(a)
            integer, intent(in) :: a
        end function f
    end interface

    procedure(f), pointer :: p
    integer :: a
    p => square
    a = p(3)
    p => cube
    a = a + p(2)
    print *, a
contains
    integer function square(a)
        integer, intent(in) :: a
        square = a * a
    end function square
    integer function cube(a)
        integer, intent(in) :: a
        cube = a * a * a
    end function cube
end program procedure_associated_pointer_calls_pointer_reassignment
"#,
    );
    assert_eq!(out, vec!["17"]);
}

#[test]
fn procedure_associated_pointer_calls_nullify_guarded_call() {
    let out = run_prints(
        r#"
program procedure_associated_pointer_calls_nullify_guarded_call
    abstract interface
        subroutine op(a, b)
            integer, intent(in) :: a, b
        end subroutine op
    end interface

    procedure(op), pointer :: p
    integer :: s

    p => add
    if (associated(p)) then
        call p(5, 6)
    end if
    s = 10
    print *, s
contains
    subroutine add(a, b)
        integer, intent(in) :: a, b
    end subroutine add
end program procedure_associated_pointer_calls_nullify_guarded_call
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn procedure_associated_pointer_calls_nesting_pointer_to_component() {
    let out = run_prints(
        r#"
program procedure_associated_pointer_calls_nesting_pointer_to_component
    type callback_holder
        procedure(scale_iface), pointer, nopass :: f
    end type callback_holder

    abstract interface
        integer function scale_iface(x, factor)
            integer, intent(in) :: x
            integer, intent(in), optional :: factor
        end function scale_iface
    end interface

    type(callback_holder) :: holder
    integer :: value

    holder%f => scale
    value = holder%f(4, 3)
    print *, value
contains
    integer function scale(x, factor)
        integer, intent(in) :: x
        integer, intent(in), optional :: factor
        if (present(factor)) then
            scale = x * factor
        else
            scale = x
        end if
    end function scale
end program procedure_associated_pointer_calls_nesting_pointer_to_component
"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn procedure_associated_pointer_calls_array_of_procedures() {
    let out = run_prints(
        r#"
program procedure_associated_pointer_calls_array_of_procedures
    abstract interface
        integer function f(a)
            integer, intent(in) :: a
        end function f
    end interface

    procedure(f), pointer :: fn_table(2)
    integer :: total

    fn_table(1) => double
    fn_table(2) => triple
    total = fn_table(1)(2) + fn_table(2)(2)
    print *, total
contains
    integer function double(a)
        integer, intent(in) :: a
        double = 2 * a
    end function double
    integer function triple(a)
        integer, intent(in) :: a
        triple = 3 * a
    end function triple
end program procedure_associated_pointer_calls_array_of_procedures
"#,
    );
    assert_eq!(out, vec!["10"]);
}
