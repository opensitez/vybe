! vybe-test: fortran/specification_part/spec_optional_13
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program t
implicit none
integer :: seen
call s()
if (seen /= 0) then
    print *, "FAIL: want [0] got [", seen, "]"
    stop 1
end if
call s(9)
if (seen /= 9) then
    print *, "FAIL: want [9] got [", seen, "]"
    stop 1
end if
contains
subroutine s(x)
implicit none
integer, optional :: x
if (present(x)) then
    seen = x
else
    seen = 0
end if
end subroutine s
end program t
