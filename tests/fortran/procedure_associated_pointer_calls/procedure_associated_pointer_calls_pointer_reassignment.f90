! vybe-test: fortran/procedure_associated_pointer_calls/procedure_associated_pointer_calls_pointer_reassignment
! origin: languages/fortran/tests/fortran/test_procedure_associated_pointer_calls.rs

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
    if ((a) /= 17) then
    print *, "FAIL: want [17] got [", a, "]"
    stop 1
end if
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
