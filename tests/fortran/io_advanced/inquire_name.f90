! vybe-test: fortran/io_advanced/inquire_name
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    character(len=100) :: fname
    open(unit=10, status='scratch')
    inquire(unit=10, name=fname)
    close(10)
    print *, 'ok'
end program test
