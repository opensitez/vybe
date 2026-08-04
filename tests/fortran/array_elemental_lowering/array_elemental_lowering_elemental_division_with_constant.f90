! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_elemental_division_with_constant
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_elemental_division_with_constant
    integer, allocatable :: values(:)
    integer, allocatable :: half(:)
    values = (/ 2, 4, 6, 8 /)
    half = values / 2
    if ((size(half)) /= 4) then
    print *, "FAIL: want [4] got [", size(half), "]"
    stop 1
end if
    if ((sum(half)) /= 10) then
    print *, "FAIL: want [10] got [", sum(half), "]"
    stop 1
end if
    if ((half(4)) /= 4) then
    print *, "FAIL: want [4] got [", half(4), "]"
    stop 1
end if
end program array_elemental_lowering_elemental_division_with_constant
