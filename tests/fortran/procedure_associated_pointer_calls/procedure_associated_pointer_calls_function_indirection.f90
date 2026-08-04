! vybe-test: fortran/procedure_associated_pointer_calls/procedure_associated_pointer_calls_function_indirection
! origin: languages/fortran/tests/fortran/test_procedure_associated_pointer_calls.rs

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
    if ((result) /= 8) then
    print *, "FAIL: want [8] got [", result, "]"
    stop 1
end if
contains
    integer function add_mul(a, b)
        integer, intent(in) :: a, b
        add_mul = a * b + a
    end function add_mul
end program procedure_associated_pointer_calls_function_indirection
