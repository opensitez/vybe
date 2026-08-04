! vybe-test: fortran/fortran_io_quality/io_quality_print_format_integer
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_print_format_integer
    integer :: value
    value = 123
    print '(I0)', value
end program io_quality_print_format_integer
