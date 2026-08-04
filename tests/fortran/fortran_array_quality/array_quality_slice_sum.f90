! vybe-test: fortran/fortran_array_quality/array_quality_slice_sum
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_slice_sum
    integer, dimension(6) :: values
    values = (/ 1, 2, 3, 4, 5, 6 /)
    if ((values(2:5:2)(1) + values(2:5:2)(2)) /= 6) then
    print *, "FAIL: want [6] got [", values(2:5:2)(1) + values(2:5:2)(2), "]"
    stop 1
end if
end program array_quality_slice_sum
