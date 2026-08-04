! vybe-test: fortran/fortran_array_quality/array_quality_two_dimensional_layout
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_two_dimensional_layout
    integer, dimension(2,3) :: mat
    mat = reshape((/1, 2, 3, 4, 5, 6/), (/2,3/))
    if ((mat(2,3)) /= 6) then
    print *, "FAIL: want [6] got [", mat(2,3), "]"
    stop 1
end if
end program array_quality_two_dimensional_layout
