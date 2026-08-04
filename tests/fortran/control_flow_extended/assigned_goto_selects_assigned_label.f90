! vybe-test: fortran/control_flow_extended/assigned_goto_selects_assigned_label
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: n, x
x = 0
assign 20 to n
go to n
x = 99
20 x = 20
if ((x) /= 20) then
    print *, "FAIL: want [20] got [", x, "]"
    stop 1
end if
end program t
