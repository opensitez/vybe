! vybe-test: fortran/forall_advanced/forall_mask_lower_triangle
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    real :: m(4,4)
    m = 0.0
    forall (i = 1:4, j = 1:4, j < i)
        m(i,j) = real(i + j)
    end forall
    print *, m(3,1)
    print *, m(1,2)
end program test
