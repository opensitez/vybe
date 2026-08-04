! vybe-test: fortran/program_units_error_recovery/program_units_error_recovery_unbalanced_if_rejected
! origin: languages/fortran/tests/fortran/test_program_units_error_recovery.rs
program program_units_error_recovery_unbalanced_if_rejected
integer :: x
if (.true.) print *, 1
end program
