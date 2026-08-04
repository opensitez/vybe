! vybe-test: fortran/program_units_error_recovery/program_units_error_recovery_invalid_return_name
! origin: languages/fortran/tests/fortran/test_program_units_error_recovery.rs
function f() result(r)
r = 1
end f
