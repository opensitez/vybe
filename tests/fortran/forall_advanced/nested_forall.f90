! vybe-test: fortran/forall_advanced/nested_forall
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: m(3,3)
    m = 0
    forall (i = 1:3)
        forall (j = 1:3)
            m(i,j) = i * 10 + j
        end forall
    end forall
    print *, m(2,3)
end program test
