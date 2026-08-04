! vybe-test: fortran/stop_error_extended/error_stop_with_quiet_clause
! origin: languages/fortran/tests/fortran/test_stop_error_extended.rs

program t
    logical :: bad = .false.
    if (bad) error stop 9, quiet = .true.
    print *, 'ok'
end program t
