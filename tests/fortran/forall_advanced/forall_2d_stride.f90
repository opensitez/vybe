! vybe-test: fortran/forall_advanced/forall_2d_stride
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: m(6,6) = 0
    forall (i = 1:6:2, j = 1:6:2)
        m(i,j) = i * j
    end forall
    print *, m(1,1)
    print *, m(3,3)
    print *, m(2,2)
end program test
