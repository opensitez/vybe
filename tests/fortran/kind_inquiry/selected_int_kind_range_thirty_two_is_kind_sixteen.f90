! vybe-test: fortran/kind_inquiry/selected_int_kind_range_thirty_two_is_kind_sixteen
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((selected_int_kind(32)) /= 16) then
    print *, "FAIL: want [16] got [", selected_int_kind(32), "]"
    stop 1
end if
end program t
