! vybe-test: fortran/stop_error_status_signaling/stop_error_status_signaling_error_stop_with_code
! origin: languages/fortran/tests/fortran/test_stop_error_status_signaling.rs
program stop_error_status_signaling_error_stop_with_code
integer :: x
x = 0
if (x /= 0) error stop 17
print *, x
end program stop_error_status_signaling_error_stop_with_code
