! vybe-test: fortran/fortran2008/spread_intrinsic
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: a(3) = [1, 2, 3]
    integer :: m(3, 4)
    m = spread(a, 2, 4)
    print *, m(2, 1)
end program test
