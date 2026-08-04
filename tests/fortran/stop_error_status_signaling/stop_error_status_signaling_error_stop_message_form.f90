! vybe-test: fortran/stop_error_status_signaling/stop_error_status_signaling_error_stop_message_form
! origin: languages/fortran/tests/fortran/test_stop_error_status_signaling.rs
program stop_error_status_signaling_error_stop_message_form
integer :: n
n = 1
if (n < 0) error stop 'invalid-negative'
print *, n
end program stop_error_status_signaling_error_stop_message_form
