! vybe-test: fortran/io_advanced/write_err
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: ios
    write(*, *, iostat=ios) 42
    if (ios /= 0) print *, 'write error'
end program test
