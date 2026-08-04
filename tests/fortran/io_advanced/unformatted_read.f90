! vybe-test: fortran/io_advanced/unformatted_read
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: n
    open(unit=10, file='bin.dat', form='unformatted', status='replace')
    write(10) 99
    rewind(10)
    read(10) n
    close(10)
    print *, n
end program test
