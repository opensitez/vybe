! vybe-test: fortran/fortran_io_quality/io_quality_repeat_format
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_repeat_format
    integer :: i
    character(len=40) :: text
    text = 'ok '
    write (text, '(I0,A)') 4, text
    print *, trim(text)
end program io_quality_repeat_format
