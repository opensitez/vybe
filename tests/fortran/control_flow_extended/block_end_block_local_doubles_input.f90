! vybe-test: fortran/control_flow_extended/block_end_block_local_doubles_input
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: x = 6
block
integer :: y
y = x * 2
if ((y) /= 12) then
    print *, "FAIL: want [12] got [", y, "]"
    stop 1
end if
end block
if ((x) /= 6) then
    print *, "FAIL: want [6] got [", x, "]"
    stop 1
end if
end program t
