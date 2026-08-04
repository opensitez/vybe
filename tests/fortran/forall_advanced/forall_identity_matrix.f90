! vybe-test: fortran/forall_advanced/forall_identity_matrix
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    real :: id(5,5)
    id = 0.0
    forall (i = 1:5)
        id(i,i) = 1.0
    end forall
    print *, id(1,1)
    print *, id(1,2)
end program test
