! vybe-test: fortran/stop_error_extended/guarded_stop_in_if_not_taken
! origin: languages/fortran/tests/fortran/test_stop_error_extended.rs
program t
logical :: ok = .true.
if (.not. ok) stop 1
if (trim('continued') /= "continued") then
    print *, "FAIL: want [continued] got [", 'continued', "]"
    stop 1
end if
end program t
