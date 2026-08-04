! vybe-test: fortran/fortran_io_quality/io_quality_print_star_default
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_print_star_default
    if ((42) /= 42) then
    print *, "FAIL: want [42] got [", 42, "]"
    stop 1
end if
end program io_quality_print_star_default
