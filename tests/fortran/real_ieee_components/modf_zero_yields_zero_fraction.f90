! vybe-test: fortran/real_ieee_components/modf_zero_yields_zero_fraction
! origin: languages/fortran/tests/fortran/test_real_ieee_components.rs
program t
real :: f
integer :: i
i = int(0.0)
f = 0.0 - real(i)
if (i /= 0) then
    print *, "FAIL: want [0] got [", i, "]"
    stop 1
end if
if (nint(f * 100) /= 0) then
    print *, "FAIL: want [0] got [", nint(f * 100), "]"
    stop 1
end if
end program t
