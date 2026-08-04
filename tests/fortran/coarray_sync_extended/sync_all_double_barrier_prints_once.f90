! vybe-test: fortran/coarray_sync_extended/sync_all_double_barrier_prints_once
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
integer :: step[*]
step = this_image()
sync all
sync all
if (this_image() == 1) print *, step
end program t
