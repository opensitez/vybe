! vybe-test: fortran/control_flow_extended/goto_inside_do_escapes_iteration_body
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: i, s
s = 0
do i = 1, 10
if (i == 4) goto 30
s = s + i
end do
30 print *, s
end program t
