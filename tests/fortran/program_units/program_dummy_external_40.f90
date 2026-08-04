! vybe-test: fortran/program_units/program_dummy_external_40
! origin: languages/fortran/tests/fortran/test_program_units.rs
subroutine apply(f)
external f
call f()
end subroutine apply
