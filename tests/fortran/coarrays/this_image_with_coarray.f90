! vybe-test: fortran/coarrays/this_image_with_coarray
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x[*]
    integer :: coindices(1)
    x = this_image()
    coindices = this_image(x, 1)
    print *, coindices(1)
end program test
