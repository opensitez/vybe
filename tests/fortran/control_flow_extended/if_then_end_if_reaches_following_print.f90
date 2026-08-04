! vybe-test: fortran/control_flow_extended/if_then_end_if_reaches_following_print
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: x = 3
if (x > 0) then
if (trim('in') /= "in") then
    print *, "FAIL: want [in] got [", 'in', "]"
    stop 1
end if
end if
if (trim('out') /= "out") then
    print *, "FAIL: want [out] got [", 'out', "]"
    stop 1
end if
end program t
