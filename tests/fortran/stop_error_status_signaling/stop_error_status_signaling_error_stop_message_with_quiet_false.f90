! vybe-test: fortran/stop_error_status_signaling/stop_error_status_signaling_error_stop_message_with_quiet_false
! origin: languages/fortran/tests/fortran/test_stop_error_status_signaling.rs
program stop_error_status_signaling_error_stop_message_with_quiet_false
integer :: flag
flag = 0
if (flag == 0) then
error stop 'status ok', quiet = .false.
else
print *, 'not triggered'
end if
print *, 'end'
end program stop_error_status_signaling_error_stop_message_with_quiet_false
