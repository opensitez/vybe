! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_implicit_type_coercion_through_real
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_implicit_type_coercion_through_real
    integer, allocatable :: ints(:)
    real, allocatable :: vals(:)
    ints = (/ 1, 2, 3, 4, 5 /)
    vals = real(ints) + 0.5
    if ((sum(ints)) /= 15) then
    print *, "FAIL: want [15] got [", sum(ints), "]"
    stop 1
end if
    if ((nint(sum(vals))) /= 17) then
    print *, "FAIL: want [17] got [", nint(sum(vals)), "]"
    stop 1
end if
    if ((nint(vals(1))) /= 2) then
    print *, "FAIL: want [2] got [", nint(vals(1)), "]"
    stop 1
end if
    if ((nint(vals(size(vals)))) /= 6) then
    print *, "FAIL: want [6] got [", nint(vals(size(vals))), "]"
    stop 1
end if
end program array_elemental_lowering_implicit_type_coercion_through_real
