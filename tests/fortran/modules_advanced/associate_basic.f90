! vybe-test: fortran/modules_advanced/associate_basic
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

program test
    type :: Point
        real :: x, y
    end type Point
    type(Point) :: p
    p%x = 3.0
    p%y = 4.0
    associate(xx => p%x, yy => p%y)
        print *, sqrt(xx*xx + yy*yy)
    end associate
end program test
