! vybe-test: fortran/kind_inquiry/selected_real_kind_matches_kind_of_double_literal
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer, parameter :: k = selected_real_kind(15, 307)
if ((k) /= 8) then
    print *, "FAIL: want [8] got [", k, "]"
    stop 1
end if
end program t
