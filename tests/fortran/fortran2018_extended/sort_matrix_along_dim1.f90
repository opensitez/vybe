! vybe-test: fortran/fortran2018_extended/sort_matrix_along_dim1
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    integer :: m(2,3) = reshape([3, 1, 4, 1, 5, 9], [2, 3])
    call sort(m, dim=1)
    print *, m(1, 1), m(2, 1)
end program t
