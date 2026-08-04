! vybe-test: fortran/arrays/slice_range
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: a(5) = [1, 2, 3, 4, 5]
    integer :: b(3)
    b = a(2:4)
    print *, b(1)
end program test
