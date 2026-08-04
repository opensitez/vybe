! vybe-test: fortran/program_units/program_keyword_args_23
! origin: languages/fortran/tests/fortran/test_program_units.rs
subroutine s(x,y)
integer :: x,y
end subroutine s
program p
call s(y=2, x=1)
end program p
