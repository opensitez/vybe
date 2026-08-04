! vybe-test: fortran/legacy/equivalence_basic
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: a
    integer :: b
    equivalence (a, b)
    a = 42
    print *, b
end program test
