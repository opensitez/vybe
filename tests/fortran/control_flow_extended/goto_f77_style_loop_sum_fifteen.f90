! vybe-test: fortran/control_flow_extended/goto_f77_style_loop_sum_fifteen
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: i, s
i = 1
s = 0
10 if (i > 5) goto 20
s = s + i
i = i + 1
goto 10
20 print *, s
end program t
