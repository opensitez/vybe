! vybe-test: fortran/kind_inquiry/selected_int_kind_range_thirty_two_unavailable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((selected_int_kind(32)) /= 8) then
    print *, "FAIL: want [8] got [", selected_int_kind(32), "]"
    stop 1
end if
end program t
