! vybe-test: fortran/forall_advanced/forall_tridiagonal
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    real :: m(5,5)
    m = 0.0
    forall (i = 1:5)
        m(i,i) = 2.0
    end forall
    forall (i = 1:4)
        m(i,i+1) = -1.0
        m(i+1,i) = -1.0
    end forall
    print *, m(1,1)
    print *, m(1,2)
end program test
