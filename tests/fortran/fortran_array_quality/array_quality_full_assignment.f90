! vybe-test: fortran/fortran_array_quality/array_quality_full_assignment
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_full_assignment
    integer, dimension(5) :: values
    values = (/ 1, 2, 3, 4, 5 /)
    if ((values(1) + values(5)) /= 6) then
    print *, "FAIL: want [6] got [", values(1) + values(5), "]"
    stop 1
end if
end program array_quality_full_assignment
