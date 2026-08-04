! vybe-test: fortran/stop_error_status_signaling/stop_error_status_signaling_error_stop_with_identifier_code
! origin: languages/fortran/tests/fortran/test_stop_error_status_signaling.rs
program stop_error_status_signaling_error_stop_with_identifier_code
integer :: status
status = 3
if (status > 0) error stop status
print *, status
end program stop_error_status_signaling_error_stop_with_identifier_code
