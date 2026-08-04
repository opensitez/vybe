! vybe-test: fortran/array_slice_lvalue_semantics/array_slice_lvalue_semantics_strided_section_assignment
! origin: languages/fortran/tests/fortran/test_array_slice_lvalue_semantics.rs

program array_slice_lvalue_semantics_strided_section_assignment
    integer :: values(1:8)
    values = (/1, 2, 3, 4, 5, 6, 7, 8/)
    values(2:8:2) = 1
    if ((values(2)) /= 1) then
    print *, "FAIL: want [1] got [", values(2), "]"
    stop 1
end if
    if ((values(4)) /= 1) then
    print *, "FAIL: want [1] got [", values(4), "]"
    stop 1
end if
    if ((values(6)) /= 1) then
    print *, "FAIL: want [1] got [", values(6), "]"
    stop 1
end if
    if ((values(8)) /= 1) then
    print *, "FAIL: want [1] got [", values(8), "]"
    stop 1
end if
    if ((sum(values)) /= 20) then
    print *, "FAIL: want [20] got [", sum(values), "]"
    stop 1
end if
end program array_slice_lvalue_semantics_strided_section_assignment
