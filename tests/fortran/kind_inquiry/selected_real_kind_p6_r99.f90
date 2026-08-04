! vybe-test: fortran/kind_inquiry/selected_real_kind_p6_r99
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((selected_real_kind(6, 99)) /= 8) then
    print *, "FAIL: want [8] got [", selected_real_kind(6, 99), "]"
    stop 1
end if
end program t
