! vybe-test: fortran/implicit_interfaces/implicit_interfaces_19
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
implicit none
call caller()
contains
subroutine caller()
integer :: x
external f
x = f()
end subroutine caller
end program p
