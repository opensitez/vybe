! vybe-test: fortran/program_units_error_recovery/program_units_error_recovery_invalid_fixed_loop_rejected
! origin: languages/fortran/tests/fortran/test_program_units_error_recovery.rs
program program_units_error_recovery_invalid_fixed_loop_rejected
do i = 1, 10
print *, i
end do
