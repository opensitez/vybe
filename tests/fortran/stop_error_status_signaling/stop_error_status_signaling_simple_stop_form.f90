! vybe-test: fortran/stop_error_status_signaling/stop_error_status_signaling_simple_stop_form
! origin: languages/fortran/tests/fortran/test_stop_error_status_signaling.rs
program stop_error_status_signaling_simple_stop_form
integer :: x
x = 1
if (x > 0) stop 0
print *, 'unreachable'
end program stop_error_status_signaling_simple_stop_form
