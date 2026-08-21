! vybe-test: fortran/findloc/findloc_basic
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: a(6) = [3, 1, 4, 1, 5, 9]
    integer :: loc(1)
    loc = findloc(a, 4)
    print *, loc(1)
end program test
