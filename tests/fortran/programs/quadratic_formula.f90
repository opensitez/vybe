! vybe-test: fortran/programs/quadratic_formula
! origin: languages/fortran/tests/fortran/test_programs.rs

program quadratic
    real :: a, b, c, discriminant, x1, x2
    a = 1.0
    b = -5.0
    c = 6.0
    discriminant = b**2 - 4.0*a*c
    if (discriminant >= 0.0) then
        x1 = (-b + sqrt(discriminant)) / (2.0*a)
        x2 = (-b - sqrt(discriminant)) / (2.0*a)
        print *, x1
        print *, x2
    end if
end program quadratic
