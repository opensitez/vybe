! vybe-test: fortran/bit_btest_ishftc/ieor_negative_one_with_one
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
if ((ieor(-1, 1)) /= -2) then
    print *, "FAIL: want [-2] got [", ieor(-1, 1), "]"
    stop 1
end if
end program t
