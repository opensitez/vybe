! vybe-test: fortran/io_advanced/open_read_write
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    open(unit=20, file='data.txt', status='replace', action='write')
    write(20, '(A)') 'test data'
    close(20)
end program test
