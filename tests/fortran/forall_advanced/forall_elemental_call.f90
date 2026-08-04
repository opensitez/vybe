! vybe-test: fortran/forall_advanced/forall_elemental_call
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    real :: a(5) = [1.0, 4.0, 9.0, 16.0, 25.0]
    real :: b(5)
    forall (i = 1:5)
        b(i) = sqrt(a(i))
    end forall
    print *, b(1)
    print *, b(4)
end program test
