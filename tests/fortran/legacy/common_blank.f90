! vybe-test: fortran/legacy/common_blank
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: a, b
    common a, b
    a = 1
    b = 2
    print *, a * b
end program test
