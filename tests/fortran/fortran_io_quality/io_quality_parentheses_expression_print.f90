! vybe-test: fortran/fortran_io_quality/io_quality_parentheses_expression_print
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_parentheses_expression_print
    if ((2 * (3 + 1)) /= 8) then
    print *, "FAIL: want [8] got [", 2 * (3 + 1), "]"
    stop 1
end if
end program io_quality_parentheses_expression_print
