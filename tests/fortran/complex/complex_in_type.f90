! vybe-test: fortran/complex/complex_in_type
! origin: languages/fortran/tests/fortran/test_complex.rs

program test
    type :: Phasor
        real :: magnitude
        complex :: value
    end type Phasor
    type(Phasor) :: p
    p%magnitude = 5.0
    p%value = (3.0, 4.0)
    print *, p%magnitude
end program test
