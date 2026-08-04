! vybe-test: fortran/program_units/program_internal_proc_03
! origin: languages/fortran/tests/fortran/test_program_units.rs
program p
call s()
contains
subroutine s()
print *, 1
end subroutine s
end program p
