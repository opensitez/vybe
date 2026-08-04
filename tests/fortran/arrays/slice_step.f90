! vybe-test: fortran/arrays/slice_step
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(6) = [10, 20, 30, 40, 50, 60]
    integer :: b(3)
    b = a(1:6:2)
    print *, b(1)
end program test
