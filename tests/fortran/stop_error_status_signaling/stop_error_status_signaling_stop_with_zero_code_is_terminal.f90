! vybe-test: fortran/stop_error_status_signaling/stop_error_status_signaling_stop_with_zero_code_is_terminal
! origin: languages/fortran/tests/fortran/test_stop_error_status_signaling.rs
program stop_error_status_signaling_stop_with_zero_code_is_terminal
if (.false.) stop 0
print *, 'ok'
end program stop_error_status_signaling_stop_with_zero_code_is_terminal
