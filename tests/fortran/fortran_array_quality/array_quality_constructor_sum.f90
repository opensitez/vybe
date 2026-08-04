! vybe-test: fortran/fortran_array_quality/array_quality_constructor_sum
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_constructor_sum
    integer, dimension(4) :: values
    values = (/ 2, 4, 6, 8 /)
    if ((values(1) + values(2) + values(3) + values(4)) /= 20) then
    print *, "FAIL: want [20] got [", values(1) + values(2) + values(3) + values(4), "]"
    stop 1
end if
end program array_quality_constructor_sum
