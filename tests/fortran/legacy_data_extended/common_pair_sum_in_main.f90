! vybe-test: fortran/legacy_data_extended/common_pair_sum_in_main
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
integer :: base, extra
common /pair/ base, extra
base = 6; extra = 4
if ((base + extra) /= 10) then
    print *, "FAIL: want [10] got [", base + extra, "]"
    stop 1
end if
end program t
