! vybe-test: fortran/bit_btest_ishftc/ibset_negative_two_bit_zero
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
if ((ibset(-2, 0)) /= -1) then
    print *, "FAIL: want [-1] got [", ibset(-2, 0), "]"
    stop 1
end if
end program t
