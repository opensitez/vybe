! vybe-test: fortran/coarray_sync_extended/sync_all_after_remote_read_pattern
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
integer :: buf[*]
buf = this_image() * 3
sync all
if (this_image() == 1) print *, buf
end program t
