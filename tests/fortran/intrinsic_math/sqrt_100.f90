! vybe-test: fortran/intrinsic_math/sqrt_100
! origin: languages/fortran/tests/fortran/test_intrinsic_math.rs
program t
if ((sqrt(100.0)) /= 10) then
    print *, "FAIL: want [10] got [", sqrt(100.0), "]"
    stop 1
end if
end program t
