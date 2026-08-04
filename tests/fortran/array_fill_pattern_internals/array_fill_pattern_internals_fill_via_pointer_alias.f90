! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_via_pointer_alias
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_via_pointer_alias
    integer, target :: base(3)
    integer, pointer :: alias(:)
    base = (/ 1, 2, 3 /)
    alias => base
    alias = 9
    alias(2:3) = 11
    if ((sum(base)) /= 31) then
    print *, "FAIL: want [31] got [", sum(base), "]"
    stop 1
end if
    if ((base(1)) /= 9) then
    print *, "FAIL: want [9] got [", base(1), "]"
    stop 1
end if
    if ((alias(3)) /= 11) then
    print *, "FAIL: want [11] got [", alias(3), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_via_pointer_alias
