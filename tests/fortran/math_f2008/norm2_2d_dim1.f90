! vybe-test: fortran/math_f2008/norm2_2d_dim1
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real :: m(2,3) = reshape([1.,2.,3.,4.,5.,6.],[2,3])
    real :: col_norms(3)
    col_norms = norm2(m, dim=1)
    print *, col_norms(1)
end program test
