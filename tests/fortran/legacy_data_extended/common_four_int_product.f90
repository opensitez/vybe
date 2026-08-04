! vybe-test: fortran/legacy_data_extended/common_four_int_product
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: a, b, c, d
common /quad/ a, b, c, d
a = 2; b = 3; c = 4; d = 5
if ((a * b * c * d) /= 120) then
    print *, "FAIL: want [120] got [", a * b * c * d, "]"
    stop 1
end if
end program t
