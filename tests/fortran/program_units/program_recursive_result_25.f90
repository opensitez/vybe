! vybe-test: fortran/program_units/program_recursive_result_25
! origin: languages/fortran/tests/fortran/test_program_units.rs
recursive integer function f(n) result(r)
integer :: n
if (n <= 1) then
 r = 1
else
 r = n * f(n-1)
end if
end function f
