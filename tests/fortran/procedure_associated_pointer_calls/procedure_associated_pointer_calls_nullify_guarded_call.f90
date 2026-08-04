! vybe-test: fortran/procedure_associated_pointer_calls/procedure_associated_pointer_calls_nullify_guarded_call
! origin: languages/fortran/tests/fortran/test_procedure_associated_pointer_calls.rs

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
    if ((s) /= 10) then
    print *, "FAIL: want [10] got [", s, "]"
    stop 1
end if
contains
    subroutine add(a, b)
        integer, intent(in) :: a, b
    end subroutine add
end program procedure_associated_pointer_calls_nullify_guarded_call
