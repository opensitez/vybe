! vybe-test: fortran/interfaces/if_target_18
! origin: languages/fortran/tests/fortran/test_interfaces.rs
program t
integer :: v
v = 1
call s(v)
if (v /= 9) then
    print *, "FAIL: want [9] got [", v, "]"
    stop 1
end if
contains
subroutine s(x)
integer,target::x
integer,pointer::p
p => x
p = 9
end subroutine s
end program t
