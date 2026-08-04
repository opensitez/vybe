! vybe-test: fortran/io_advanced/open_newunit
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: u
    open(newunit=u, file='tmp.txt', status='replace')
    write(u, *) 'newunit test'
    close(u)
end program test
