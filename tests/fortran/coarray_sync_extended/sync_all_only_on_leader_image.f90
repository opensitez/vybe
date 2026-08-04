! vybe-test: fortran/coarray_sync_extended/sync_all_only_on_leader_image
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs

program t
    integer :: flag[*]
    flag = 0
    if (this_image() == 1) sync all
    flag = this_image()
    print *, flag
end program t
