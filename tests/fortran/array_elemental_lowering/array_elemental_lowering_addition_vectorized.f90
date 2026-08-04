! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_addition_vectorized
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_addition_vectorized
    integer, allocatable :: left(:), right(:)
    left = (/ 1, 2, 3, 4 /)
    right = (/ 4, 3, 2, 1 /)
    if ((sum(left + right)) /= 20) then
    print *, "FAIL: want [20] got [", sum(left + right), "]"
    stop 1
end if
    if (((left + right)(1)) /= 5) then
    print *, "FAIL: want [5] got [", (left + right)(1), "]"
    stop 1
end if
    if (((left + right)(4)) /= 5) then
    print *, "FAIL: want [5] got [", (left + right)(4), "]"
    stop 1
end if
end program array_elemental_lowering_addition_vectorized
