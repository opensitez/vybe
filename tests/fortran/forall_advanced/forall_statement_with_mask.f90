! vybe-test: fortran/forall_advanced/forall_statement_with_mask
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: a(10) = 0
    forall (i = 1:10, mod(i,3) == 0) a(i) = i
    print *, a(3)
    print *, a(6)
    print *, a(4)
end program test
