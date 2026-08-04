! vybe-test: fortran/array_slice_lvalue_semantics/array_slice_lvalue_semantics_direct_assignment_updates_elements
! origin: languages/fortran/tests/fortran/test_array_slice_lvalue_semantics.rs

program array_slice_lvalue_semantics_direct_assignment_updates_elements
    integer :: values(1:6)
    values = (/1, 2, 3, 4, 5, 6/)
    values(2:5) = 0
    if ((sum(values)) /= 7) then
    print *, "FAIL: want [7] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
    if ((values(6)) /= 6) then
    print *, "FAIL: want [6] got [", values(6), "]"
    stop 1
end if
end program array_slice_lvalue_semantics_direct_assignment_updates_elements
