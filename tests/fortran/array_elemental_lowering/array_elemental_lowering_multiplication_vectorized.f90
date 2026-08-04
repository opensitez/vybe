! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_multiplication_vectorized
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_multiplication_vectorized
    integer, allocatable :: left(:), right(:)
    left = (/ 2, 3, 4, 5 /)
    right = (/ 1, 2, 3, 4 /)
    if ((sum(left * right)) /= 40) then
    print *, "FAIL: want [40] got [", sum(left * right), "]"
    stop 1
end if
    if (((left * right)(1)) /= 2) then
    print *, "FAIL: want [2] got [", (left * right)(1), "]"
    stop 1
end if
    if (((left * right)(3)) /= 12) then
    print *, "FAIL: want [12] got [", (left * right)(3), "]"
    stop 1
end if
end program array_elemental_lowering_multiplication_vectorized
