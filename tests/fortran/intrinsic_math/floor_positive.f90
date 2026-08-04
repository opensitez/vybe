! vybe-test: fortran/intrinsic_math/floor_positive
! origin: languages/fortran/tests/fortran/test_intrinsic_math.rs
program t
if ((floor(3.7)) /= 3) then
    print *, "FAIL: want [3] got [", floor(3.7), "]"
    stop 1
end if
end program t
