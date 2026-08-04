! vybe-test: fortran/program_units/program_dummy_proc_17
! origin: languages/fortran/tests/fortran/test_program_units.rs
subroutine apply(f)
external f
call f()
end subroutine apply
