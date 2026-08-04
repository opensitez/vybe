! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_all_and_any_maskes
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_all_and_any_maskes
    integer, allocatable :: values(:)
    values = (/ 1, 1, 1, 0 /)
    if ((merge(1, 0, all(values == 1))) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, all(values == 1)), "]"
    stop 1
end if
    if ((merge(1, 0, any(values == 0))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, any(values == 0)), "]"
    stop 1
end if
    if ((count(values /= 1)) /= 1) then
    print *, "FAIL: want [1] got [", count(values /= 1), "]"
    stop 1
end if
end program array_elemental_lowering_all_and_any_maskes
