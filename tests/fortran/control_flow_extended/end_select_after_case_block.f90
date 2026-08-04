! vybe-test: fortran/control_flow_extended/end_select_after_case_block
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: v = 2, r
select case (v)
case (1)
r = 10
case (2)
r = 20
case default
r = 99
end select
if ((r) /= 20) then
    print *, "FAIL: want [20] got [", r, "]"
    stop 1
end if
end program t
