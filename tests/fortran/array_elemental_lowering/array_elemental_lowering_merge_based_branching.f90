! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_merge_based_branching
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_merge_based_branching
    integer, allocatable :: values(:)
    integer, allocatable :: merged(:)
    values = (/ -1, 2, -3, 4 /)
    merged = merge(10, 0, values > 0)
    if ((sum(merged)) /= 20) then
    print *, "FAIL: want [20] got [", sum(merged), "]"
    stop 1
end if
    if ((merged(1)) /= 0) then
    print *, "FAIL: want [0] got [", merged(1), "]"
    stop 1
end if
    if ((merged(4)) /= 10) then
    print *, "FAIL: want [10] got [", merged(4), "]"
    stop 1
end if
end program array_elemental_lowering_merge_based_branching
