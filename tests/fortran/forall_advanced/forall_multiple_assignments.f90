! vybe-test: fortran/forall_advanced/forall_multiple_assignments
! origin: languages/fortran/tests/fortran/test_forall_advanced.rs

program test
    real :: x(5), y(5), z(5)
    x = [1.0, 2.0, 3.0, 4.0, 5.0]
    y = [5.0, 4.0, 3.0, 2.0, 1.0]
    forall (i = 1:5)
        z(i) = x(i) + y(i)
        x(i) = x(i) * 2.0
    end forall
    print *, z(1)
    print *, x(1)
end program test
