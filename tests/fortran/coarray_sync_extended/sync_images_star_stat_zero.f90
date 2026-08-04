! vybe-test: fortran/coarray_sync_extended/sync_images_star_stat_zero
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
integer :: stat
sync images (*, stat=stat)
if ((stat) /= 0) then
    print *, "FAIL: want [0] got [", stat, "]"
    stop 1
end if
end program t
