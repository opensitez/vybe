! vybe-test: fortran/legacy_data_extended/save_initial_value_single_call
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
call show()
contains
subroutine show()
integer, save :: tally = 17
if ((tally) /= 17) then
    print *, "FAIL: want [17] got [", tally, "]"
    stop 1
end if
end subroutine show
end program t
