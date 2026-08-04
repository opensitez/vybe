! vybe-test: fortran/legacy_data_extended/common_append_second_declaration
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: a, b, c
common /grp/ a, b
common /grp/ c
a = 1; b = 2; c = 3
if ((a + b + c) /= 6) then
    print *, "FAIL: want [6] got [", a + b + c, "]"
    stop 1
end if
end program t
