! vybe-test: fortran/stop_error_extended/error_stop_expression
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: code = 0
    if (code /= 0) error stop code
    print *, 'ok'
end program test
