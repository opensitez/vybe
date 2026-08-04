! vybe-test: fortran/optional_arguments/optional_arguments_05
! origin: languages/fortran/tests/fortran/test_optional_arguments.rs
program p
interface
subroutine s(x)
integer, optional :: x
end subroutine s
end interface
call s()
end program p
