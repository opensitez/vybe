! vybe-test: fortran/legacy/equivalence_array_scalar
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: arr(4)
    integer :: first
    equivalence (arr(1), first)
    arr(1) = 99
    print *, first
end program test
