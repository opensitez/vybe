! vybe-test: fortran/fortran_io_quality/io_quality_character_output_trim
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_character_output_trim
    character(len=12) :: word
    word = 'fortran'
    print '(A)', trim(word)
end program io_quality_character_output_trim
