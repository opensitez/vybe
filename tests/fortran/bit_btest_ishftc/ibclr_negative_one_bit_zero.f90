! vybe-test: fortran/bit_btest_ishftc/ibclr_negative_one_bit_zero
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
if ((ibclr(-1, 0)) /= -2) then
    print *, "FAIL: want [-2] got [", ibclr(-1, 0), "]"
    stop 1
end if
end program t
