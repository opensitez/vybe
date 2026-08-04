! vybe-test: fortran/legacy_data_extended/common_seven_var_sum
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: s(7)
common /wide/ s
s = [(i, i = 1, 7)]
if ((sum(s)) /= 28) then
    print *, "FAIL: want [28] got [", sum(s), "]"
    stop 1
end if
end program t
