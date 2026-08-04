! vybe-test: fortran/floating_point_negative_zero_cases/test_floating_point_negative_zero_cases_preserves_sign
! origin: languages/fortran/tests/fortran/test_floating_point_negative_zero_cases.rs

program test_floating_point_negative_zero_cases
    real :: x
    x = -0.0
    if ((sign(1.0, x)) /= -1) then
    print *, "FAIL: want [-1] got [", sign(1.0, x), "]"
    stop 1
end if
    if ((x == 0.0) .neqv. .true.) then
    print *, "FAIL: want [True] got [", x == 0.0, "]"
    stop 1
end if
end program test_floating_point_negative_zero_cases
