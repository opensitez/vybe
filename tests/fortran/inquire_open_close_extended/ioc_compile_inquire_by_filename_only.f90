! vybe-test: fortran/inquire_open_close_extended/ioc_compile_inquire_by_filename_only
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs

program t
    logical :: opened
    inquire(file='ioc_ext_fn.dat', opened=opened)
    print *, 0
end program t
