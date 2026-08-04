! vybe-test: fortran/forall_advanced/forall_statement_form
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    real :: a(5)
    forall (i = 1:5) a(i) = real(i) ** 2
    print *, a(3)
end program test
