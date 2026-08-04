! vybe-test: fortran/fortran2018/stop_variable
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: code = 0
    print *, 'ok'
    stop code
end program test
