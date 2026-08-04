! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_abs_on_signed_integers
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_abs_on_signed_integers
    integer, allocatable :: values(:)
    values = (/ -5, -1, 0, 2, -3 /)
    if ((sum(abs(values))) /= 11) then
    print *, "FAIL: want [11] got [", sum(abs(values)), "]"
    stop 1
end if
    if ((abs(values(1))) /= 5) then
    print *, "FAIL: want [5] got [", abs(values(1)), "]"
    stop 1
end if
    if ((abs(values(5))) /= 3) then
    print *, "FAIL: want [3] got [", abs(values(5)), "]"
    stop 1
end if
end program array_elemental_lowering_abs_on_signed_integers
