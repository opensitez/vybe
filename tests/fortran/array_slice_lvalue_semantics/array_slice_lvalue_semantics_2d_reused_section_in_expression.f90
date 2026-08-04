! vybe-test: fortran/array_slice_lvalue_semantics/array_slice_lvalue_semantics_2d_reused_section_in_expression
! origin: languages/fortran/tests/fortran/test_array_slice_lvalue_semantics.rs

program array_slice_lvalue_semantics_2d_reused_section_in_expression
    integer :: values(3, 3)
    values = reshape((/1, 2, 3, 4, 5, 6, 7, 8, 9/), (/3, 3/))
    values(2:3, :) = values(1:2, :) + 1
    if ((values(2, 1)) /= 2) then
    print *, "FAIL: want [2] got [", values(2, 1), "]"
    stop 1
end if
    if ((values(3, 3)) /= 8) then
    print *, "FAIL: want [8] got [", values(3, 3), "]"
    stop 1
end if
    if ((sum(values)) /= 45) then
    print *, "FAIL: want [45] got [", sum(values), "]"
    stop 1
end if
end program array_slice_lvalue_semantics_2d_reused_section_in_expression
