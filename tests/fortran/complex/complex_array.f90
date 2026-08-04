! vybe-test: fortran/complex/complex_array
! origin: languages/fortran/tests/fortran/test_complex.rs

program test
    complex :: v(3) = [(0.0,0.0), (1.0,0.0), (0.0,1.0)]
    print *, real(v(2))
end program test
