! vybe-test: fortran/control_flow_extended/computed_goto_index_out_of_range_falls_through
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: n = 4
go to (10, 20), n
if (trim('fallthrough') /= "fallthrough") then
    print *, "FAIL: want [fallthrough] got [", 'fallthrough', "]"
    stop 1
end if
10 print *, 'one'
20 print *, 'two'
end program t
