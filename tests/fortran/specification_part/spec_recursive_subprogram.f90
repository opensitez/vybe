! vybe-test: fortran/specification_part/spec_recursive_subprogram
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program t
implicit none
integer :: fac
if (fac(5) /= 120) then
    print *, "FAIL: want [120] got [", fac(5), "]"
    stop 1
end if
end program t
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
