! vybe-test: fortran/legacy/continue_label
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: i
    do i = 1, 5
        if (mod(i, 2) == 0) goto 100
        print *, i
100     continue
    end do
end program test
