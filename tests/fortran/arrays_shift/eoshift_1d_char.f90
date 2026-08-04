! vybe-test: fortran/arrays_shift/eoshift_1d_char
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    character(len=1) :: a(4) = ['a', 'b', 'c', 'd']
    character(len=1) :: b(4)
    b = eoshift(a, 1, boundary=' ')
    print *, b(4)
end program test
