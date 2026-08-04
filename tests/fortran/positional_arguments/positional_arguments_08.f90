! vybe-test: fortran/positional_arguments/positional_arguments_08
! origin: languages/fortran/tests/fortran/test_positional_arguments.rs
subroutine s(a)
integer::a(2)
end
program p
integer::a(2)=[1,2]
call s(a)
end program p
