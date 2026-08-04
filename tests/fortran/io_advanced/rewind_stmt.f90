! vybe-test: fortran/io_advanced/rewind_stmt
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    open(unit=10, status='scratch')
    write(10, *) 1, 2, 3
    rewind(10)
    close(10)
end program test
