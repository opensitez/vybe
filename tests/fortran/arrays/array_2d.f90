! vybe-test: fortran/arrays/array_2d
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: m(3,3)
    integer :: i, j
    do i = 1, 3
        do j = 1, 3
            m(i,j) = i * 3 + j
        end do
    end do
    print *, m(2,2)
end program test
