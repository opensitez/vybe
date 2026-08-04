! vybe-test: fortran/io_advanced/open_status_old
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    open(unit=30, file='existing.txt', status='old', action='read', iostat=ios)
    integer :: ios
    if (ios /= 0) print *, 'file not found'
end program test
