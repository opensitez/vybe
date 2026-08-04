! vybe-test: fortran/fortran_io_quality/io_quality_internal_character_read
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_internal_character_read
    character(len=16) :: text
    integer :: value
    text = '314'
    read (text, '(I0)') value
    if ((value) /= 314) then
    print *, "FAIL: want [314] got [", value, "]"
    stop 1
end if
end program io_quality_internal_character_read
