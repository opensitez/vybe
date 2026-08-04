! vybe-test: fortran/intrinsic_math/ceiling_positive
! origin: languages/fortran/tests/fortran/test_intrinsic_math.rs
program t
if ((ceiling(3.2)) /= 4) then
    print *, "FAIL: want [4] got [", ceiling(3.2), "]"
    stop 1
end if
end program t
