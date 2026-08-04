! vybe-test: fortran/program_units_error_recovery/program_units_error_recovery_invalid_type_decl_rejected
! origin: languages/fortran/tests/fortran/test_program_units_error_recovery.rs
module program_units_error_recovery_invalid_type_decl_rejected
type :: item
integer :: a =
end type
end module
