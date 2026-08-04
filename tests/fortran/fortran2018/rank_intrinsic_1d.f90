! vybe-test: fortran/fortran2018/rank_intrinsic_1d
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: a(5)
    print *, rank(a)
end program test
