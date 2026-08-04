! vybe-test: fortran/complex/complex_array_ops
! origin: languages/fortran/tests/fortran/test_complex.rs

program test
    complex :: a(3) = [(1.0,0.0), (2.0,0.0), (3.0,0.0)]
    complex :: b(3) = [(0.0,1.0), (0.0,2.0), (0.0,3.0)]
    complex :: c(3)
    c = a + b
    print *, real(c(1))
end program test
