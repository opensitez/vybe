! vybe-test: fortran/coarrays/coarray_remote_read
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x[*]
    x = this_image() * 10
    sync all
    if (this_image() == 1 .and. num_images() >= 2) then
        print *, x[2]
    end if
end program test
