! vybe-test: fortran/forall_advanced/forall_outer_product
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    real :: u(4) = [1.0, 2.0, 3.0, 4.0]
    real :: v(4) = [1.0, 2.0, 3.0, 4.0]
    real :: m(4,4)
    forall (i = 1:4, j = 1:4)
        m(i,j) = u(i) * v(j)
    end forall
    print *, m(2,3)
end program test
