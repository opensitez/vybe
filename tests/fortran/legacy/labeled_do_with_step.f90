! vybe-test: fortran/legacy/labeled_do_with_step
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: i, s
    s = 0
    do 10 i = 0, 10, 2
        s = s + i
10  continue
    print *, s
end program test
