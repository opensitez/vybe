! vybe-test: fortran/fortran_io_quality/io_quality_float_scientific_format
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_float_scientific_format
    real :: value
    value = 1.25
    print '(E10.3)', value
end program io_quality_float_scientific_format
