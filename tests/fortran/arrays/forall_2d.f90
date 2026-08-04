! vybe-test: fortran/arrays/forall_2d
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    real :: m(3,3)
    forall (i = 1:3, j = 1:3)
        m(i,j) = real(i) + real(j)
    end forall
    print *, m(1,1)
end program test
