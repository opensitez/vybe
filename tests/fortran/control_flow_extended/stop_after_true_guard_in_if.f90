! vybe-test: fortran/control_flow_extended/stop_after_true_guard_in_if
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: code = 1
if (code /= 0) stop 0
print *, 'run'
end program t
