! vybe-test: fortran/stop_error_extended/guarded_stop_in_if_taken_halts
! origin: languages/fortran/tests/fortran/test_stop_error_extended.rs
program t
integer :: code = 1
if (code /= 0) stop 0
print *, 'run'
end program t
