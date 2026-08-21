! vybe-test: fortran/rank_intrinsic/rank_intrinsic_3d
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    real :: m(2,3,4)
    print *, rank(m)
end program test
