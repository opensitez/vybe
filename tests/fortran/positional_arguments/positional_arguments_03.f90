! vybe-test: fortran/positional_arguments/positional_arguments_03
! origin: languages/fortran/tests/fortran/test_positional_arguments.rs
subroutine s(x,y,z)
integer::x,y,z
end
program p
call s(1,2,3)
end program p
