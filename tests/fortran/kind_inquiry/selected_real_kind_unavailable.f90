! vybe-test: fortran/kind_inquiry/selected_real_kind_unavailable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((selected_real_kind(999)) /= -1) then
    print *, "FAIL: want [-1] got [", selected_real_kind(999), "]"
    stop 1
end if
end program t
