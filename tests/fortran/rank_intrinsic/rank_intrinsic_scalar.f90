! vybe-test: fortran/rank_intrinsic/rank_intrinsic_scalar
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: x = 5
    print *, rank(x)
end program test
