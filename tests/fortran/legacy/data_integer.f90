! vybe-test: fortran/legacy/data_integer
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: x, y
    data x /42/, y /99/
    print *, x + y
end program test
