! vybe-test: fortran/stop_error_status_signaling/stop_error_status_signaling_simple_stop_form_runtime
! origin: languages/fortran/tests/fortran/test_stop_error_status_signaling.rs
program stop_error_status_signaling_simple_stop_form_runtime
integer :: x
x = 0
if (x > 0) stop 0
if ((x) /= 0) then
    print *, "FAIL: want [0] got [", x, "]"
    stop 1
end if
end program stop_error_status_signaling_simple_stop_form_runtime
