! vybe-test: fortran/explicit_interfaces/explicit_interfaces_02
! origin: languages/fortran/tests/fortran/test_explicit_interfaces.rs
subroutine s(x)
integer::x
x = x + 1
end subroutine s
program t

interface
subroutine s(x)
integer::x
end subroutine s
end interface
integer :: v
v = 4
call s(v)
if (v /= 5) then
    print *, "FAIL: want [5] got [", v, "]"
    stop 1
end if
end program t
