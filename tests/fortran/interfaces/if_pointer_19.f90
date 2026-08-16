! vybe-test: fortran/interfaces/if_pointer_19
! origin: languages/fortran/tests/fortran/test_interfaces.rs
program t
integer, pointer :: p
integer, target :: v
v = 1
p => v
call s(p)
if (v /= 7) then
    print *, "FAIL: want [7] got [", v, "]"
    stop 1
end if
contains
subroutine s(x)
integer,pointer::x
x = 7
end subroutine s
end program t
