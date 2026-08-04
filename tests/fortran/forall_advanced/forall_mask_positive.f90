! vybe-test: fortran/forall_advanced/forall_mask_positive
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: a(10) = [(i - 5, i=1,10)]
    integer :: b(10)
    b = 0
    forall (i = 1:10, a(i) > 0)
        b(i) = a(i)
    end forall
    print *, b(8)
    print *, b(3)
end program test
