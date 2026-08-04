! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_bit_mask_style_boolean_transform
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_bit_mask_style_boolean_transform
    integer, allocatable :: values(:)
    integer, allocatable :: flags(:)
    values = (/ 1, 2, 4, 8 /)
    flags = iand(values, 2)
    if ((sum(flags)) /= 10) then
    print *, "FAIL: want [10] got [", sum(flags), "]"
    stop 1
end if
    if ((flags(1)) /= 0) then
    print *, "FAIL: want [0] got [", flags(1), "]"
    stop 1
end if
    if ((flags(4)) /= 0) then
    print *, "FAIL: want [0] got [", flags(4), "]"
    stop 1
end if
    if ((count(flags > 0)) /= 1) then
    print *, "FAIL: want [1] got [", count(flags > 0), "]"
    stop 1
end if
end program array_elemental_lowering_bit_mask_style_boolean_transform
