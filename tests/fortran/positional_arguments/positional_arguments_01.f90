! vybe-test: fortran/positional_arguments/positional_arguments_01
! origin: languages/fortran/tests/fortran/test_positional_arguments.rs
subroutine s(x)
integer::x
end
program p
call s(1)
end program p
