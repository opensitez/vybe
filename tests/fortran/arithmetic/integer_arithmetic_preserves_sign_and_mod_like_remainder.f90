! vybe-test: fortran/arithmetic/integer_arithmetic_preserves_sign_and_mod_like_remainder
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((7 / 2) /= 3) then
    print *, "FAIL: want [3] got [", 7 / 2, "]"
    stop 1
end if
if ((5 - 2 * 2) /= 1) then
    print *, "FAIL: want [1] got [", 5 - 2 * 2, "]"
    stop 1
end if
end program t
