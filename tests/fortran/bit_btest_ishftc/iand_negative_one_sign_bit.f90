! vybe-test: fortran/bit_btest_ishftc/iand_negative_one_sign_bit
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
if ((iand(-1, ishft(1, 31))) /= -2147483648) then
    print *, "FAIL: want [-2147483648] got [", iand(-1, ishft(1, 31)), "]"
    stop 1
end if
end program t
