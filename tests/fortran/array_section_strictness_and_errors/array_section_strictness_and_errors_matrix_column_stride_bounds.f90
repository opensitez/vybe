! vybe-test: fortran/array_section_strictness_and_errors/array_section_strictness_and_errors_matrix_column_stride_bounds
! origin: languages/fortran/tests/fortran/test_array_section_strictness_and_errors.rs

program array_section_strictness_and_errors_matrix_column_stride_bounds
    integer :: values(4, 4)
    values = reshape((/ (i, i = 1, 16) /), (/4, 4/))
    if ((size(values(:, 2:4:2), 1)) /= 4) then
    print *, "FAIL: want [4] got [", size(values(:, 2:4:2), 1), "]"
    stop 1
end if
    if ((size(values(:, 2:4:2), 2)) /= 2) then
    print *, "FAIL: want [2] got [", size(values(:, 2:4:2), 2), "]"
    stop 1
end if
    if ((sum(values(:, 2:4:2))) /= 30) then
    print *, "FAIL: want [30] got [", sum(values(:, 2:4:2)), "]"
    stop 1
end if
end program array_section_strictness_and_errors_matrix_column_stride_bounds
