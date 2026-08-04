! vybe-test: fortran/forall_advanced/forall_mask_even
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: a(10)
    a = 0
    forall (i = 1:10, mod(i, 2) == 0)
        a(i) = i
    end forall
    print *, a(4)
    print *, a(3)
end program test
