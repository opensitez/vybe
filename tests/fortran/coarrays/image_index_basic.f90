! vybe-test: fortran/coarrays/image_index_basic
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x[2,*]
    integer :: sub(2) = [1, 1]
    print *, image_index(x, sub)
end program test
