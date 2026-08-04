! vybe-test: fortran/arithmetic_extended/ishft_half_max_left_wraps_sign
! origin: languages/fortran/tests/fortran/test_arithmetic_extended.rs
program t
if ((ishft(1073741824, 1)) /= -2147483648) then
    print *, "FAIL: want [-2147483648] got [", ishft(1073741824, 1), "]"
    stop 1
end if
end program t
