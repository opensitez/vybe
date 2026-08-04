! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_reduction_after_cast
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_reduction_after_cast
    real, allocatable :: values(:)
    integer :: count_hi
    values = (/ 1.2, 2.8, 3.6, 4.4 /)
    count_hi = count(int(values) >= 3)
    if ((count_hi) /= 2) then
    print *, "FAIL: want [2] got [", count_hi, "]"
    stop 1
end if
    if ((nint(maxval(values))) /= 4) then
    print *, "FAIL: want [4] got [", nint(maxval(values)), "]"
    stop 1
end if
    if ((nint(minval(values))) /= 1) then
    print *, "FAIL: want [1] got [", nint(minval(values)), "]"
    stop 1
end if
end program array_elemental_lowering_reduction_after_cast
