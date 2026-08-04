! vybe-test: fortran/arithmetic/integer_division_truncates_toward_zero
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((7 / 2) /= 3) then
    print *, "FAIL: want [3] got [", 7 / 2, "]"
    stop 1
end if
if ((-17 / 5) /= -3) then
    print *, "FAIL: want [-3] got [", -17 / 5, "]"
    stop 1
end if
end program t
