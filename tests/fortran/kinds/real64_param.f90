! vybe-test: fortran/kinds/real64_param
! origin: languages/fortran/tests/fortran/test_kinds.rs

program test
    integer, parameter :: real64 = 8
    real(kind=real64) :: x = 1.0_8
    print *, x
end program test
