! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_sectionwise_copy_like_transform
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_sectionwise_copy_like_transform
    integer :: source(1:6)
    integer :: target(1:6)
    source = (/ 1, 2, 3, 4, 5, 6 /)
    target = 0
    target(2:5) = source(2:5)
    if ((source(1)) /= 1) then
    print *, "FAIL: want [1] got [", source(1), "]"
    stop 1
end if
    if ((target(1)) /= 0) then
    print *, "FAIL: want [0] got [", target(1), "]"
    stop 1
end if
    if ((sum(target)) /= 14) then
    print *, "FAIL: want [14] got [", sum(target), "]"
    stop 1
end if
    if ((target(2) + target(5)) /= 7) then
    print *, "FAIL: want [7] got [", target(2) + target(5), "]"
    stop 1
end if
end program array_elemental_lowering_sectionwise_copy_like_transform
