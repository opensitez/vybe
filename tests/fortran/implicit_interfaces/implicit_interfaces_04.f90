! vybe-test: fortran/implicit_interfaces/implicit_interfaces_04
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
external s
integer :: x
x=1
call s(x)
end program p
