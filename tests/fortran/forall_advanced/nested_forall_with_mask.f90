! vybe-test: fortran/forall_advanced/nested_forall_with_mask
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: m(4,4) = 0
    forall (i = 1:4, i <= 2)
        forall (j = 1:4, j > i)
            m(i,j) = i + j
        end forall
    end forall
    print *, m(1,2)
    print *, m(1,1)
end program test
