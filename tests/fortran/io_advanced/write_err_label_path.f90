! vybe-test: fortran/io_advanced/write_err_label_path
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: ios
    write(999, *, err=10, iostat=ios) 42
    print *, 0
10 continue
    if (ios /= 0) print *, 1
end program test
