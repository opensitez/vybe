! vybe-test: fortran/inquire_open_close_extended/ioc_compile_close_err_label
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs

program t
    integer :: ios
    open(10, status='scratch')
    close(10, iostat=ios)
    print *, ios
end program t
