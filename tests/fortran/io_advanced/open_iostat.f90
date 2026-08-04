! vybe-test: fortran/io_advanced/open_iostat
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: ios
    open(unit=10, file='nosuchfile.txt', status='old', iostat=ios)
    if (ios /= 0) print *, 'could not open'
end program test
