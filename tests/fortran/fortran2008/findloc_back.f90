! vybe-test: fortran/fortran2008/findloc_back
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: a(6) = [1, 2, 1, 2, 1, 2]
    integer :: loc(1)
    loc = findloc(a, 1, back=.true.)
    print *, loc(1)
end program test
