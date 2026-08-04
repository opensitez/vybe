! vybe-test: fortran/legacy/data_repeated
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: a(6)
    data a /6*0/
    print *, a(1)
end program test
