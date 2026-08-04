! vybe-test: fortran/legacy/goto_basic
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: x = 0
    goto 10
    x = 999
10  continue
    print *, x
end program test
