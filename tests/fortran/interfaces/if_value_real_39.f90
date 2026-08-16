! vybe-test: fortran/interfaces/if_value_real_39
! origin: languages/fortran/tests/fortran/test_interfaces.rs
program t
real :: v
v = 1.5
call s(v)
if (abs(v - 1.5) > 1.0e-6) then
    print *, "FAIL: want [1.5] got [", v, "]"
    stop 1
end if
contains
subroutine s(x)
real, value :: x
x = x + 1.0
end subroutine s
end program t
