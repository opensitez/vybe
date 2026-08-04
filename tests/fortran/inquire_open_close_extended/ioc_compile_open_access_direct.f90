! vybe-test: fortran/inquire_open_close_extended/ioc_compile_open_access_direct
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs

program t
    open(10, file='ioc_ext_direct.dat', access='direct', recl=10, status='replace')
    write(10, rec=1) 'test'
    close(10, status='delete')
    print *, 1
end program t
