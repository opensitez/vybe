! vybe-test: fortran/intrinsic_math/sqrt_25
! origin: languages/fortran/tests/fortran/test_intrinsic_math.rs
program t
if ((sqrt(25.0)) /= 5) then
    print *, "FAIL: want [5] got [", sqrt(25.0), "]"
    stop 1
end if
end program t
