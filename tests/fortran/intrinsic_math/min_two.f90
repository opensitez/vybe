! vybe-test: fortran/intrinsic_math/min_two
! origin: languages/fortran/tests/fortran/test_intrinsic_math.rs
program t
if ((min(3, 7)) /= 3) then
    print *, "FAIL: want [3] got [", min(3, 7), "]"
    stop 1
end if
end program t
