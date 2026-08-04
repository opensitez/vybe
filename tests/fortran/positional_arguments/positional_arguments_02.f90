! vybe-test: fortran/positional_arguments/positional_arguments_02
! origin: languages/fortran/tests/fortran/test_positional_arguments.rs
subroutine s(x,y)
integer::x,y
end
program p
call s(1,2)
end program p
