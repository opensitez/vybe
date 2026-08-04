! vybe-test: fortran/modules_advanced/associate_expr
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

program test
    real :: x = 3.0, y = 4.0
    associate(hyp => sqrt(x*x + y*y))
        print *, hyp
    end associate
end program test
