! vybe-test: fortran/positional_arguments/positional_arguments_04
! origin: languages/fortran/tests/fortran/test_positional_arguments.rs
subroutine s(x)
real::x
end
program p
call s(1.0)
end program p
