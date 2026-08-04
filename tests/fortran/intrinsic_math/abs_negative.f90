! vybe-test: fortran/intrinsic_math/abs_negative
! origin: languages/fortran/tests/fortran/test_intrinsic_math.rs
program t
if ((abs(-42)) /= 42) then
    print *, "FAIL: want [42] got [", abs(-42), "]"
    stop 1
end if
end program t
