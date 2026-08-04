! vybe-test: fortran/io_advanced/inquire_exist
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    logical :: exists
    inquire(file='test.txt', exist=exists)
    print *, exists
end program test
