! vybe-test: fortran/inquire_open_close_extended/ioc_compile_open_delim_quote
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs

program t
    open(10, file='ioc_ext_delim.dat', status='replace', delim='quote')
    write(10, *) 'a'
    close(10, status='delete')
    print *, 1
end program t
