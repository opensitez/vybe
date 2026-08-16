! vybe-test: fortran/interfaces/if_value_16
! origin: languages/fortran/tests/fortran/test_interfaces.rs
program t
integer :: v
v = 3
call s(v)
if (v /= 3) then
    print *, "FAIL: want [3] got [", v, "]"
    stop 1
end if
contains
subroutine s(x)
integer,value::x
x = x + 1
end subroutine s
end program t
