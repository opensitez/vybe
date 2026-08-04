! vybe-test: fortran/procedure_associated_pointer_calls/procedure_associated_pointer_calls_basic_subroutine_indirection
! origin: languages/fortran/tests/fortran/test_procedure_associated_pointer_calls.rs

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
    if ((result) /= 7) then
    print *, "FAIL: want [7] got [", result, "]"
    stop 1
end if
contains
    subroutine add_two(a, b, c)
        integer, intent(in) :: a, b
        integer, intent(out) :: c
        c = a + b
    end subroutine add_two
end program procedure_associated_pointer_calls_basic_subroutine_indirection
