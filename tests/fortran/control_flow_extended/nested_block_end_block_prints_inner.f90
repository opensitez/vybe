! vybe-test: fortran/control_flow_extended/nested_block_end_block_prints_inner
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: a = 2
block
integer :: b
b = a + 3
block
integer :: c
c = b + 5
if ((c) /= 10) then
    print *, "FAIL: want [10] got [", c, "]"
    stop 1
end if
end block
end block
if ((a) /= 2) then
    print *, "FAIL: want [2] got [", a, "]"
    stop 1
end if
end program t
