! vybe-test: fortran/fortran_io_quality/io_quality_labelled_write_field_width
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_labelled_write_field_width
    integer :: value
    value = 42
    write (*, '(I4)') value
end program io_quality_labelled_write_field_width
