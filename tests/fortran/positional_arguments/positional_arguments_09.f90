! vybe-test: fortran/positional_arguments/positional_arguments_09
! origin: languages/fortran/tests/fortran/test_positional_arguments.rs
subroutine s(x,y)
integer::x
real::y
end
program p
call s(1,2.0)
end program p
