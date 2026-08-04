! vybe-test: fortran/specification_part/spec_recursive_subprogram
! origin: languages/fortran/tests/fortran/test_specification_part.rs
recursive function fac(n) result(r)
implicit none
integer, intent(in) :: n
integer :: r
if (n <= 1) then
 r = 1
else
 r = n * fac(n - 1)
end if
end function fac
