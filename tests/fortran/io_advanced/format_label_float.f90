! vybe-test: fortran/io_advanced/format_label_float
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    real :: x = 2.718
200 format(F10.4)
    write(*, 200) x
end program test
