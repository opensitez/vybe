! vybe-test: fortran/specification_part/spec_value_23
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program t
implicit none
integer :: v
v = 2
call s(v)
if (v /= 2) then
    print *, "FAIL: want [2] got [", v, "]"
    stop 1
end if
contains
subroutine s(x)
implicit none
integer, value :: x
x = x + 100
end subroutine s
end program t
