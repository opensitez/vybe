! vybe-test: fortran/array_slice_lvalue_semantics/array_slice_lvalue_semantics_section_as_ghost_target
! origin: languages/fortran/tests/fortran/test_array_slice_lvalue_semantics.rs

program array_slice_lvalue_semantics_section_as_ghost_target
    integer :: values(1:6)
    values = (/1, 2, 3, 4, 5, 6/)
    values(4:) = 0
    if ((values(3)) /= 3) then
    print *, "FAIL: want [3] got [", values(3), "]"
    stop 1
end if
    if ((values(4)) /= 0) then
    print *, "FAIL: want [0] got [", values(4), "]"
    stop 1
end if
    if ((values(6)) /= 0) then
    print *, "FAIL: want [0] got [", values(6), "]"
    stop 1
end if
    if ((sum(values)) /= 6) then
    print *, "FAIL: want [6] got [", sum(values), "]"
    stop 1
end if
end program array_slice_lvalue_semantics_section_as_ghost_target
