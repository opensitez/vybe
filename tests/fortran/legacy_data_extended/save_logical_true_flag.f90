! vybe-test: fortran/legacy_data_extended/save_logical_true_flag
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
call flag_once()
contains
subroutine flag_once()
logical, save :: on = .true.
if ((on) .neqv. .true.) then
    print *, "FAIL: want [true] got [", on, "]"
    stop 1
end if
end subroutine flag_once
end program t
