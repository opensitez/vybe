! vybe-test: fortran/positional_arguments/positional_arguments_07
! origin: languages/fortran/tests/fortran/test_positional_arguments.rs
subroutine s(x)
complex::x
end
program p
call s((1.0,2.0))
end program p
