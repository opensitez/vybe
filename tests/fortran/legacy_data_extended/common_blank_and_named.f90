! vybe-test: fortran/legacy_data_extended/common_blank_and_named
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: a, b
real :: r
common a, b
common /one/ r
a = 3; b = 4; r = 0.5
if ((a + b) /= 7) then
    print *, "FAIL: want [7] got [", a + b, "]"
    stop 1
end if
if (abs((r) - 0.5) > 1.0e-6) then
    print *, "FAIL: want [0.5] got [", r, "]"
    stop 1
end if
end program t
