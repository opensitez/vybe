! vybe-test: fortran/array_slice_lvalue_semantics/array_slice_lvalue_semantics_mixed_stride_reassign
! origin: languages/fortran/tests/fortran/test_array_slice_lvalue_semantics.rs

program array_slice_lvalue_semantics_mixed_stride_reassign
    integer :: values(1:10)
    values = (/1, 1, 1, 1, 1, 1, 1, 1, 1, 1/)
    values(2:10:3) = (/2, 3, 4/)
    if ((values(2)) /= 2) then
    print *, "FAIL: want [2] got [", values(2), "]"
    stop 1
end if
    if ((values(5)) /= 3) then
    print *, "FAIL: want [3] got [", values(5), "]"
    stop 1
end if
    if ((values(8)) /= 4) then
    print *, "FAIL: want [4] got [", values(8), "]"
    stop 1
end if
    if ((sum(values)) /= 16) then
    print *, "FAIL: want [16] got [", sum(values), "]"
    stop 1
end if
end program array_slice_lvalue_semantics_mixed_stride_reassign
