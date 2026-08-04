! vybe-test: fortran/array_slice_lvalue_semantics/array_slice_lvalue_semantics_matrix_subsection_increment
! origin: languages/fortran/tests/fortran/test_array_slice_lvalue_semantics.rs

program array_slice_lvalue_semantics_matrix_subsection_increment
    integer :: values(3, 3)
    values = reshape((/1, 2, 3, 4, 5, 6, 7, 8, 9/), (/3, 3/))
    values(2:3, 2:3) = values(2:3, 2:3) + 1
    if ((values(2, 2)) /= 6) then
    print *, "FAIL: want [6] got [", values(2, 2), "]"
    stop 1
end if
    if ((values(3, 3)) /= 10) then
    print *, "FAIL: want [10] got [", values(3, 3), "]"
    stop 1
end if
    if ((sum(values)) /= 53) then
    print *, "FAIL: want [53] got [", sum(values), "]"
    stop 1
end if
end program array_slice_lvalue_semantics_matrix_subsection_increment
