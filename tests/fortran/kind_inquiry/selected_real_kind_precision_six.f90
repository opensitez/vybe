! vybe-test: fortran/kind_inquiry/selected_real_kind_precision_six
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((selected_real_kind(6)) /= 4) then
    print *, "FAIL: want [4] got [", selected_real_kind(6), "]"
    stop 1
end if
end program t
