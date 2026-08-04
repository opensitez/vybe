! vybe-test: fortran/fortran_io_quality/io_quality_multi_value_formatting
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_multi_value_formatting
    integer :: a
    integer :: b
    integer :: c
    a = 1
    b = 2
    c = 3
    print '(I0,1x,I0,1x,I0)', a, b, c
end program io_quality_multi_value_formatting
