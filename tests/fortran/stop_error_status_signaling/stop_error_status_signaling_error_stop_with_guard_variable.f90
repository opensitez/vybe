! vybe-test: fortran/stop_error_status_signaling/stop_error_status_signaling_error_stop_with_guard_variable
! origin: languages/fortran/tests/fortran/test_stop_error_status_signaling.rs
program stop_error_status_signaling_error_stop_with_guard_variable
logical :: do_stop
do_stop = .false.
if (do_stop) error stop 99
print *, 1
end program stop_error_status_signaling_error_stop_with_guard_variable
