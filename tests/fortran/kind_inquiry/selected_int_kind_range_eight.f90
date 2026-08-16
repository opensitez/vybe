! vybe-test: fortran/kind_inquiry/selected_int_kind_range_eight
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((selected_int_kind(8)) /= 4) then
    print *, "FAIL: want [4] got [", selected_int_kind(8), "]"
    stop 1
end if
end program t
