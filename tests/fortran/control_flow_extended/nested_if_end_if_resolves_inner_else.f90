! vybe-test: fortran/control_flow_extended/nested_if_end_if_resolves_inner_else
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: a = 1, b = 0, r
if (a == 1) then
if (b == 1) then
r = 10
else
r = 20
end if
else
r = 30
end if
if ((r) /= 20) then
    print *, "FAIL: want [20] got [", r, "]"
    stop 1
end if
end program t
