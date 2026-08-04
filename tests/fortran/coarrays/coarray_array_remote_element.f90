! vybe-test: fortran/coarrays/coarray_array_remote_element
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: a(5)[*]
    a = this_image()
    sync all
    if (this_image() == 1 .and. num_images() >= 2) then
        print *, a(3)[2]
    end if
end program test
