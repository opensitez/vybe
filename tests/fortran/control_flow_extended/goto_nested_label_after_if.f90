! vybe-test: fortran/control_flow_extended/goto_nested_label_after_if
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: x
x = 0
if (x == 0) goto 10
x = 7
10 print *, x
end program t
