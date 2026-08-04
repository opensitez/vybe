! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_modular_remainder_vector
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_modular_remainder_vector
    integer, allocatable :: values(:)
    integer, allocatable :: rems(:)
    values = (/ 8, 9, 10, 11, 12 /)
    rems = mod(values, 5)
    if ((size(rems)) /= 5) then
    print *, "FAIL: want [5] got [", size(rems), "]"
    stop 1
end if
    if ((sum(rems)) /= 10) then
    print *, "FAIL: want [10] got [", sum(rems), "]"
    stop 1
end if
    if ((rems(2)) /= 4) then
    print *, "FAIL: want [4] got [", rems(2), "]"
    stop 1
end if
    if ((rems(5)) /= 2) then
    print *, "FAIL: want [2] got [", rems(5), "]"
    stop 1
end if
end program array_elemental_lowering_modular_remainder_vector
