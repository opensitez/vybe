! vybe-test: fortran/legacy_data_extended/common_named_six_total
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: i1, i2, i3, i4, i5, i6
common /pool/ i1, i2, i3, i4, i5, i6
i1 = 1; i2 = 2; i3 = 3; i4 = 4; i5 = 5; i6 = 6
if ((i1 + i2 + i3 + i4 + i5 + i6) /= 21) then
    print *, "FAIL: want [21] got [", i1 + i2 + i3 + i4 + i5 + i6, "]"
    stop 1
end if
end program t
