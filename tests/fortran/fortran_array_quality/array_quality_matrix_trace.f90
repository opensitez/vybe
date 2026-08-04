! vybe-test: fortran/fortran_array_quality/array_quality_matrix_trace
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_matrix_trace
    integer, dimension(3,3) :: m
    integer :: total
    m = reshape((/1, 0, 2, 0, 2, 0, 3, 0, 3/), (/3,3/))
    total = m(1,1) + m(2,2) + m(3,3)
    if ((total) /= 6) then
    print *, "FAIL: want [6] got [", total, "]"
    stop 1
end if
end program array_quality_matrix_trace
