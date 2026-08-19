! vybe-test: fortran/optional_arguments/optional_arguments_07
! origin: languages/fortran/tests/fortran/test_optional_arguments.rs
program p
interface
subroutine s(x,y)
integer, optional :: x,y
end subroutine s
end interface
call s(y=2)
end program p

subroutine s(x,y)
integer, optional :: x,y
end subroutine s
