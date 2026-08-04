! vybe-test: fortran/inquire_open_close_extended/ioc_compile_open_convert_native
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs

program t
    open(10, file='ioc_ext_conv.dat', status='replace', convert='native')
    write(10, '(I0)') 1
    close(10, status='delete')
    print *, 1
end program t
