! vybe-test: fortran/legacy/labeled_do_nested
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: i, j, s
    s = 0
    do 200 i = 1, 3
        do 100 j = 1, 3
            s = s + 1
100     continue
200 continue
    print *, s
end program test
