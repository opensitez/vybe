! vybe-test: fortran/kind_inquiry/selected_real_kind_p7_r37
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((selected_real_kind(7, 37)) /= 8) then
    print *, "FAIL: want [8] got [", selected_real_kind(7, 37), "]"
    stop 1
end if
end program t
