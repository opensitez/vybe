! vybe-test: fortran/positional_arguments/positional_arguments_05
! origin: languages/fortran/tests/fortran/test_positional_arguments.rs
subroutine s(x)
character(len=*)::x
end
program p
call s('abc')
end program p
