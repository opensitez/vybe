! vybe-test: fortran/forall_advanced/forall_mask_upper_triangle
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: m(4,4)
    m = 0
    forall (i = 1:4, j = 1:4, j > i)
        m(i,j) = i * 10 + j
    end forall
    print *, m(1,2)
    print *, m(1,1)
    print *, m(2,4)
end program test
