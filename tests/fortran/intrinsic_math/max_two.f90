! vybe-test: fortran/intrinsic_math/max_two
! origin: languages/fortran/tests/fortran/test_intrinsic_math.rs
program t
if ((max(3, 7)) /= 7) then
    print *, "FAIL: want [7] got [", max(3, 7), "]"
    stop 1
end if
end program t
