! vybe-test: fortran/forall_advanced/forall_mask_diagonal
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: m(5,5)
    m = 0
    forall (i = 1:5, j = 1:5, i == j)
        m(i,j) = 1
    end forall
    print *, m(1,1)
    print *, m(2,2)
    print *, m(1,2)
end program test
