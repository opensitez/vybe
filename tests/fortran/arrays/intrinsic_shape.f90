! vybe-test: fortran/arrays/intrinsic_shape
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(4)
    integer :: s(1)
    s = shape(a)
    print *, s(1)
end program test
