! vybe-test: fortran/program_units_error_recovery/program_units_error_recovery_invalid_keyword_sequence
! origin: languages/fortran/tests/fortran/test_program_units_error_recovery.rs
program p
if (.true.) then
print *, 1
else
print *, 2
end program p
