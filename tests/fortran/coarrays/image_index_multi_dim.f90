! vybe-test: fortran/coarrays/image_index_multi_dim
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x[3,4,*]
    integer :: sub(3) = [2, 3, 1]
    print *, image_index(x, sub)
end program test
