! vybe-test: fortran/positional_arguments/positional_arguments_06
! origin: languages/fortran/tests/fortran/test_positional_arguments.rs
subroutine s(x)
logical::x
end
program p
call s(.true.)
end program p
