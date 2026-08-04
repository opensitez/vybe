! vybe-test: fortran/forall_advanced/forall_mask_odd
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: a(6) = 0
    forall (i = 1:6, mod(i, 2) /= 0)
        a(i) = i * i
    end forall
    print *, a(1)
    print *, a(3)
    print *, a(2)
end program test
