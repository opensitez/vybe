! vybe-test: fortran/array_slice_lvalue_semantics/array_slice_lvalue_semantics_subsection_copy_by_shape
! origin: languages/fortran/tests/fortran/test_array_slice_lvalue_semantics.rs

program array_slice_lvalue_semantics_subsection_copy_by_shape
    integer :: source(1:4)
    integer :: target(2, 2)
    source = (/10, 11, 12, 13/)
    target = reshape(source, (/2, 2/))
    target(1, :) = 0
    if ((target(1, 1)) /= 0) then
    print *, "FAIL: want [0] got [", target(1, 1), "]"
    stop 1
end if
    if ((target(1, 2)) /= 0) then
    print *, "FAIL: want [0] got [", target(1, 2), "]"
    stop 1
end if
    if ((target(2, 1)) /= 11) then
    print *, "FAIL: want [11] got [", target(2, 1), "]"
    stop 1
end if
    if ((target(2, 2)) /= 13) then
    print *, "FAIL: want [13] got [", target(2, 2), "]"
    stop 1
end if
end program array_slice_lvalue_semantics_subsection_copy_by_shape
