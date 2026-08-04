! vybe-test: fortran/legacy_data_extended/common_blank_five_product
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: v(5)
common v
v(1) = 2; v(2) = 3; v(3) = 1; v(4) = 4; v(5) = 5
if ((v(1) * v(2) * v(3)) /= 6) then
    print *, "FAIL: want [6] got [", v(1) * v(2) * v(3), "]"
    stop 1
end if
end program t
