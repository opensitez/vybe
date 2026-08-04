! vybe-test: fortran/fortran_io_quality/io_quality_string_concatenation_print
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_string_concatenation_print
    character(len=20) :: left
    character(len=20) :: right
    left = 'foo'
    right = 'bar'
    if (trim(trim(left // right)) /= "foobar") then
    print *, "FAIL: want [foobar] got [", trim(left // right), "]"
    stop 1
end if
end program io_quality_string_concatenation_print
