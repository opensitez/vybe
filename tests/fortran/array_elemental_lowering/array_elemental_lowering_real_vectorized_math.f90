! vybe-test: fortran/array_elemental_lowering/array_elemental_lowering_real_vectorized_math
! origin: languages/fortran/tests/fortran/test_array_elemental_lowering.rs

program array_elemental_lowering_real_vectorized_math
    real, allocatable :: values(:)
    integer :: total
    values = (/ 0.5, 1.0, 1.5, 2.0 /)
    total = nint(sum(sin(values) + cos(values)))
    if ((total) /= 3) then
    print *, "FAIL: want [3] got [", total, "]"
    stop 1
end if
    if ((nint(sum(values * 2.0))) /= 8) then
    print *, "FAIL: want [8] got [", nint(sum(values * 2.0)), "]"
    stop 1
end if
end program array_elemental_lowering_real_vectorized_math
