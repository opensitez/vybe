! vybe-test: fortran/implicit_interfaces/implicit_interfaces_16
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
integer, dimension(3) :: arr
external s
call s(arr)
end program p
