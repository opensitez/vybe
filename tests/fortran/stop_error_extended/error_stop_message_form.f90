! vybe-test: fortran/stop_error_extended/error_stop_message_form
! origin: languages/fortran/tests/fortran/test_stop_error_extended.rs

program t
    logical :: bad = .false.
    if (bad) error stop 'fatal condition'
    print *, 'clear'
end program t
