! vybe-test: fortran/io_advanced/open_close_basic
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: unit = 10
    open(unit=10, file='test.txt', status='replace')
    write(10, *) 'hello'
    close(10)
end program test
