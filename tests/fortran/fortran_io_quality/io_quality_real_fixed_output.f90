! vybe-test: fortran/fortran_io_quality/io_quality_real_fixed_output
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_real_fixed_output
    real :: pi
    pi = 3.14
    print '(F6.2)', pi
end program io_quality_real_fixed_output
