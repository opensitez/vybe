! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_elemental_shift_equivalence
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_elemental_shift_equivalence
    integer, allocatable :: values(:)
    integer, allocatable :: copied(:)
    values = (/ 2, 4, 6, 8 /)
    copied = 2 * values
    if ((sum(copied)) /= 40) then
    print *, "FAIL: want [40] got [", sum(copied), "]"
    stop 1
end if
    if ((copied(2) / values(1)) /= 4) then
    print *, "FAIL: want [4] got [", copied(2) / values(1), "]"
    stop 1
end if
    if ((copied(4) - values(4)) /= 8) then
    print *, "FAIL: want [8] got [", copied(4) - values(4), "]"
    stop 1
end if
end program array_elemental_lowering_elemental_shift_equivalence
