! vybe-test: fortran/array_slice_lvalue_semantics/array_slice_lvalue_semantics_row_column_split_update
! origin: languages/fortran/tests/fortran/test_array_slice_lvalue_semantics.rs

program array_slice_lvalue_semantics_row_column_split_update
    integer :: values(4, 4)
    values = reshape((/ (i, i = 1, 16) /), (/4, 4/))
    values(1, 2:4) = 9
    values(2:4, 1) = 7
    if ((values(1, 2)) /= 9) then
    print *, "FAIL: want [9] got [", values(1, 2), "]"
    stop 1
end if
    if ((values(1, 4)) /= 9) then
    print *, "FAIL: want [9] got [", values(1, 4), "]"
    stop 1
end if
    if ((values(4, 1)) /= 7) then
    print *, "FAIL: want [7] got [", values(4, 1), "]"
    stop 1
end if
    if ((sum(values)) /= 148) then
    print *, "FAIL: want [148] got [", sum(values), "]"
    stop 1
end if
end program array_slice_lvalue_semantics_row_column_split_update
