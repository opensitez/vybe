! vybe-test: fortran/fortran_array_quality/array_quality_section_reverse_access
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_section_reverse_access
    integer, dimension(5) :: values
    values = (/ 5, 4, 3, 2, 1 /)
    print *, values(1), values(5)
end program array_quality_section_reverse_access
