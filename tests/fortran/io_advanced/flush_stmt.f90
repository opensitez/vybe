! vybe-test: fortran/io_advanced/flush_stmt
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    open(unit=10, file='out.txt', status='replace')
    write(10, *) 'buffered data'
    flush(10)
    close(10)
end program test
