! vybe-test: fortran/stop_error_extended/error_stop_variable_code
! origin: languages/fortran/tests/fortran/test_stop_error_extended.rs

program t
    integer :: code = 0
    if (code /= 0) error stop code
    print *, 'ok'
end program t
