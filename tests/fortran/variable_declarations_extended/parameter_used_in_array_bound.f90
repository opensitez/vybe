! vybe-test: fortran/variable_declarations_extended/parameter_used_in_array_bound
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, parameter :: rows = 2, cols = 2
integer, dimension(rows, cols) :: mat
mat(2, 1) = 8
if ((mat(2, 1)) /= 8) then
    print *, "FAIL: want [8] got [", mat(2, 1), "]"
    stop 1
end if
end program t
