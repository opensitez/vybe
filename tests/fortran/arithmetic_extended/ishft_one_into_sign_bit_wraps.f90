! vybe-test: fortran/arithmetic_extended/ishft_one_into_sign_bit_wraps
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((ishft(1, 31)) /= -2147483648) then
    print *, "FAIL: want [-2147483648] got [", ishft(1, 31), "]"
    stop 1
end if
end program t
