! vybe-test: fortran/arrays_shift/eoshift_2d_row_shift_array
! origin: languages/fortran/tests/fortran/test_arrays_shift.rs

program test
    integer :: m(3,3) = reshape([1,2,3,4,5,6,7,8,9],[3,3])
    integer :: shifts(3) = [0, 1, 2]
    integer :: n(3,3)
    n = eoshift(m, shifts, dim=2)
    print *, n(1,1)
end program test
