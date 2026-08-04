! vybe-test: fortran/forall_advanced/forall_with_abs
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    integer :: a(6) = [-3, 2, -1, 4, -5, 0]
    integer :: b(6)
    forall (i = 1:6)
        b(i) = abs(a(i))
    end forall
    print *, b(1)
    print *, b(2)
end program test
