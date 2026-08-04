! vybe-test: fortran/legacy/common_basic
! origin: languages/fortran/tests/fortran/test_legacy.rs

program test
    integer :: x, y
    common /data/ x, y
    x = 10
    y = 20
    print *, x + y
end program test
