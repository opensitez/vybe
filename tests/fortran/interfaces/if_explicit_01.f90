! vybe-test: fortran/interfaces/if_explicit_01
! origin: languages/fortran/tests/fortran/test_interfaces.rs
subroutine s(x)
integer :: x
x = x + 1
end subroutine s
program t
interface
subroutine s(x)
integer :: x
end subroutine s
end interface
integer :: v
v = 1
call s(v)
if (v /= 2) then
    print *, "FAIL: want [2] got [", v, "]"
    stop 1
end if
end program t
