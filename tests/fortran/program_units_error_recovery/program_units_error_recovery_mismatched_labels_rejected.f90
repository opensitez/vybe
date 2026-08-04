! vybe-test: fortran/program_units_error_recovery/program_units_error_recovery_mismatched_labels_rejected
! origin: languages/fortran/tests/fortran/test_program_units_error_recovery.rs
program p
10 print *, 1
20 print *, 2
go to 30
end program
