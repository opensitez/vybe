! vybe-test: fortran/io_advanced/inquire_size
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: fsize
    inquire(file='test.txt', size=fsize)
    print *, fsize
end program test
