! vybe-test: fortran/inquire_open_close_extended/ioc_compile_open_pad_yes
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs

program t
    open(10, file='ioc_ext_pad.dat', status='replace', pad='yes')
    write(10, '(A)') 'x'
    close(10, status='delete')
    print *, 1
end program t
