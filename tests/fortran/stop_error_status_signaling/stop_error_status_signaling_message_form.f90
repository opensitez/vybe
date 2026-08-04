! vybe-test: fortran/stop_error_status_signaling/stop_error_status_signaling_message_form
! origin: languages/fortran/tests/fortran/test_stop_error_status_signaling.rs
program stop_error_status_signaling_message_form
logical :: failed
failed = .false.
if (failed) stop 'all-good'
print *, 'running'
end program stop_error_status_signaling_message_form
