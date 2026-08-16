! vybe-test: fortran/bit_btest_ishftc/ishft_negative_two_right_one
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
if ((ishft(-2, -1)) /= 2147483647) then
    print *, "FAIL: want [2147483647] got [", ishft(-2, -1), "]"
    stop 1
end if
end program t
