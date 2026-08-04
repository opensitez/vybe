! vybe-test: fortran/procedure_associated_pointer_calls/procedure_associated_pointer_calls_array_of_procedures
! origin: languages/fortran/tests/fortran/test_procedure_associated_pointer_calls.rs

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
    if ((total) /= 10) then
    print *, "FAIL: want [10] got [", total, "]"
    stop 1
end if
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
