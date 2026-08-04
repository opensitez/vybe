! vybe-test: fortran/legacy_data_extended/common_blank_four_sum
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: a, b, c, d
common a, b, c, d
a = 1; b = 2; c = 3; d = 4
if ((a + b + c + d) /= 10) then
    print *, "FAIL: want [10] got [", a + b + c + d, "]"
    stop 1
end if
end program t
