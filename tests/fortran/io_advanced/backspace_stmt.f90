! vybe-test: fortran/io_advanced/backspace_stmt
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    open(unit=10, status='scratch')
    write(10, *) 1
    write(10, *) 2
    backspace(10)
    close(10)
end program test
