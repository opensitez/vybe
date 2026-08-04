! vybe-test: fortran/fortran_array_quality/array_quality_element_replacement
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_element_replacement
    integer, dimension(4) :: values
    integer :: i
    values = (/ 9, 4, 1, 7 /)
    do i = 1, 4
        values(i) = values(i) + 1
    end do
    print *, values(1), values(4)
end program array_quality_element_replacement
