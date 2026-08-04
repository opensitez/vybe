! vybe-test: fortran/io_advanced/flush_stdout
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    print *, 'about to flush'
    flush(6)
end program test
