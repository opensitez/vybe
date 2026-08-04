! vybe-test: fortran/legacy/labeled_do_basic
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: i, s
    s = 0
    do 100 i = 1, 5
        s = s + i
100 continue
    print *, s
end program test
