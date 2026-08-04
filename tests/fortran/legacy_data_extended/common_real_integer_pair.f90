! vybe-test: fortran/legacy_data_extended/common_real_integer_pair
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: n
real :: x
common /mix/ n, x
n = 8
x = 1.5
if ((n) /= 8) then
    print *, "FAIL: want [8] got [", n, "]"
    stop 1
end if
if (abs((x) - 1.5) > 1.0e-6) then
    print *, "FAIL: want [1.5] got [", x, "]"
    stop 1
end if
end program t
