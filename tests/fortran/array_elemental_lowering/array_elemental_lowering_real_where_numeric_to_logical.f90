! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_real_where_numeric_to_logical
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_real_where_numeric_to_logical
    real, allocatable :: values(:)
    logical, allocatable :: flags(:)
    values = (/ -1.0, 0.5, 2.25, -0.2 /)
    flags = values > 0.0
    if ((size(flags)) /= 4) then
    print *, "FAIL: want [4] got [", size(flags), "]"
    stop 1
end if
    if ((count(flags)) /= 2) then
    print *, "FAIL: want [2] got [", count(flags), "]"
    stop 1
end if
    if ((merge(1, 0, all(flags(1:2)))) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, all(flags(1:2))), "]"
    stop 1
end if
    if ((merge(1, 0, any(flags))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, any(flags)), "]"
    stop 1
end if
    if ((merge(1, 0, flags(3))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, flags(3)), "]"
    stop 1
end if
end program array_elemental_lowering_real_where_numeric_to_logical
