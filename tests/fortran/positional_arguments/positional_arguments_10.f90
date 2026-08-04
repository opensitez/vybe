! vybe-test: fortran/positional_arguments/positional_arguments_10
! origin: languages/fortran/tests/fortran/test_positional_arguments.rs
subroutine s(x,y)
character(len=*)::x
integer::y
end
program p
call s('a',1)
end program p
