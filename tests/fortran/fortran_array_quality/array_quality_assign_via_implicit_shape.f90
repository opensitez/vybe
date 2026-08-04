! vybe-test: fortran/fortran_array_quality/array_quality_assign_via_implicit_shape
! origin: languages/fortran/tests/fortran/test_fortran_array_quality.rs

program array_quality_assign_via_implicit_shape
    integer, dimension(:), allocatable :: values
    allocate(values(4))
    values = (/ 7, 8, 9, 10 /)
    if ((values(3)) /= 9) then
    print *, "FAIL: want [9] got [", values(3), "]"
    stop 1
end if
    deallocate(values)
end program array_quality_assign_via_implicit_shape
