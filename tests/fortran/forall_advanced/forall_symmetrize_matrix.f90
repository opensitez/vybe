! vybe-test: fortran/forall_advanced/forall_symmetrize_matrix
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    real :: m(3,3)
    m = 0.0
    m(1,2) = 5.0; m(1,3) = 7.0; m(2,3) = 9.0
    forall (i = 1:3, j = 1:3, i /= j)
        m(i,j) = m(i,j) + m(j,i)
    end forall
    print *, m(2,1)
end program test
