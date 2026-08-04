! vybe-test: fortran/fortran2008/findloc_not_found
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: loc(1)
    loc = findloc(a, 99)
    print *, loc(1)
end program test
