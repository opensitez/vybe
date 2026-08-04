! vybe-test: fortran/implicit_interfaces/implicit_interfaces_17
! origin: languages/fortran/tests/fortran/test_implicit_interfaces.rs
program p
integer :: i, j
external mix
call mix(i, j, i + j)
end program p
