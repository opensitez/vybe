! vybe-test: fortran/legacy/data_implied_do
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: a(5)
    data (a(i), i=1,5) /1, 2, 3, 4, 5/
    print *, a(3)
end program test
