! vybe-test: fortran/stop_error_status_signaling/stop_error_status_signaling_nested_stop_paths
! origin: languages/fortran/tests/fortran/test_stop_error_status_signaling.rs
program stop_error_status_signaling_nested_stop_paths
integer :: i
do i = 1, 2
if (i == 3) stop 1
end do
print *, 'done'
end program stop_error_status_signaling_nested_stop_paths
