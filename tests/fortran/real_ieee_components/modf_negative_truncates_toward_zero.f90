! vybe-test: fortran/real_ieee_components/modf_negative_truncates_toward_zero
! origin: languages/fortran/tests/fortran/test_real_ieee_components.rs
program t
real :: f
integer :: i
i = int(-3.75)
f = -3.75 - real(i)
if (i /= -3) then
    print *, "FAIL: want [-3] got [", i, "]"
    stop 1
end if
if (nint(f * 100) /= -75) then
    print *, "FAIL: want [-75] got [", nint(f * 100), "]"
    stop 1
end if
end program t
