! vybe-test: fortran/io_advanced/unformatted_write
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: a, b
    open(unit=10, file='bin.dat', form='unformatted', status='replace')
    write(10) 42, 99
    rewind(10)
    read(10) a, b
    close(10)
    print *, a + b
end program test
