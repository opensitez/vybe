! vybe-test: fortran/legacy/equivalence_array_scalar_runtime
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: arr(4)
    integer :: first
    equivalence (arr(1), first)
    arr(1) = 99
    if ((first) /= 99) then
    print *, "FAIL: want [99] got [", first, "]"
    stop 1
end if
end program test
