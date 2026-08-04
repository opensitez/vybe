! vybe-test: fortran/kind_inquiry/selected_int_kind_matches_kind_of_literal
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer, parameter :: k = selected_int_kind(9)
if ((k) /= 8) then
    print *, "FAIL: want [8] got [", k, "]"
    stop 1
end if
end program t
