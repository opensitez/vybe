! vybe-test: fortran/floating_point_rounding_modes/test_floating_point_rounding_modes_nint
! origin: languages/fortran/tests/fortran/test_floating_point_rounding_modes.rs

program test_floating_point_rounding_modes
    real :: value
    value = 2.6
    if ((nint(value)) /= 3) then
    print *, "FAIL: want [3] got [", nint(value), "]"
    stop 1
end if
    value = -2.6
    if ((nint(value)) /= -3) then
    print *, "FAIL: want [-3] got [", nint(value), "]"
    stop 1
end if
end program test_floating_point_rounding_modes
