! vybe-test: fortran/math_f2008/hypot_unit_circle
! origin: languages/fortran/tests/fortran/test_math_f2008.rs

program test
    real, parameter :: pi = 3.14159265
    real :: angle = pi / 4.0
    real :: h
    h = hypot(cos(angle), sin(angle))
    print *, h
end program test
