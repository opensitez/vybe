! vybe-test: fortran/coarrays/coarray_remote_write
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: shared[*]
    shared = 0
    sync all
    if (this_image() == 1) then
        shared[1] = 99
    end if
    sync all
    if (this_image() == 1) print *, shared
end program test
