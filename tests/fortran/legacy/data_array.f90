! vybe-test: fortran/legacy/data_array
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: a(5)
    data a /1, 2, 3, 4, 5/
    print *, a(3)
end program test
