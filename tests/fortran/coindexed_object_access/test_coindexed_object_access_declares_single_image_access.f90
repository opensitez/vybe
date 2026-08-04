! vybe-test: fortran/coindexed_object_access/test_coindexed_object_access_declares_single_image_access
! origin: languages/fortran/tests/fortran/test_coindexed_object_access.rs

program test_coindexed_object_access
    integer, allocatable :: shared(:)
    integer :: this
    allocate(shared(1))
    shared(1) = 7
    this = this_image()
    if ((shared(1)) /= 7) then
    print *, "FAIL: want [7] got [", shared(1), "]"
    stop 1
end if
    if ((this) /= 1) then
    print *, "FAIL: want [1] got [", this, "]"
    stop 1
end if
end program test_coindexed_object_access
