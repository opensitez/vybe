! vybe-test: fortran/stop_error_extended/stop_zero
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    print *, 'before'
    stop 0
end program test
