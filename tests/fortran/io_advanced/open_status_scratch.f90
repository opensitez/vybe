! vybe-test: fortran/io_advanced/open_status_scratch
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    open(unit=40, status='scratch')
    write(40, *) 42
    rewind(40)
    close(40, status='delete')
end program test
