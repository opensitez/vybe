! vybe-test: fortran/fortran_io_quality/io_quality_internal_character_write
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_internal_character_write
    character(len=24) :: text
    integer :: value
    value = 77
    write (text, '(I0)') value
    print *, trim(text)
end program io_quality_internal_character_write
