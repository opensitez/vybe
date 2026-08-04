! vybe-test: fortran/io_advanced/endfile_stmt
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    open(unit=10, status='scratch')
    write(10, *) 42
    endfile(10)
    close(10)
end program test
