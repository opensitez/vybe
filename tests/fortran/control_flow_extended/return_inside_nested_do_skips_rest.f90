! vybe-test: fortran/control_flow_extended/return_inside_nested_do_skips_rest
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: i
integer :: s = 0
i = 0
do while (i < 5)
i = i + 1
s = s + i
if (i == 2) return
end do
print *, 'after'
print *, s
end program t
