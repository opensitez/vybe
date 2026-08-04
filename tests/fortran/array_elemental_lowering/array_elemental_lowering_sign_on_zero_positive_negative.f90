! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_sign_on_zero_positive_negative
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_sign_on_zero_positive_negative
    integer, allocatable :: values(:)
    integer :: positive
    integer :: negative
    values = (/ -4, 0, 7 /)
    positive = sum(sign(1, values))
    negative = count(values < 0)
    if ((positive) /= 1) then
    print *, "FAIL: want [1] got [", positive, "]"
    stop 1
end if
    if ((negative) /= 1) then
    print *, "FAIL: want [1] got [", negative, "]"
    stop 1
end if
    if ((sign(-1, values(3))) /= -1) then
    print *, "FAIL: want [-1] got [", sign(-1, values(3)), "]"
    stop 1
end if
end program array_elemental_lowering_sign_on_zero_positive_negative
