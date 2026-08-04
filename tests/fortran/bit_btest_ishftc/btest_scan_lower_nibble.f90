! vybe-test: fortran/bit_btest_ishftc/btest_scan_lower_nibble
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs

program t
    integer :: i, x = 10
    do i = 0, 3
        print *, btest(x, i)
    end do
end program t
