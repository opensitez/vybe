! vybe-test: fortran/modules_advanced/module_multiple_use
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

module constants
    real, parameter :: PI = 3.14159
end module constants

module helpers
    implicit none
contains
    function double(x) result(r)
        real, intent(in) :: x
        real :: r
        r = x * 2.0
    end function double
end module helpers

program test
    use constants
    use helpers
    print *, double(PI)
end program test
