! vybe-test: fortran/fortran2008/impure_elemental_function
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: a(3) = [1, 2, 3]
    integer :: b(3)
    b = logged_double(a)
    print *, b(1)
contains
    impure elemental function logged_double(x) result(r)
        integer, intent(in) :: x
        integer :: r
        r = x * 2
    end function logged_double
end program test
