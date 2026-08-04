! vybe-test: fortran/control_flow_extended/computed_goto_selects_second_label
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: n = 2
go to (10, 20, 30), n
10 print *, 'one'; goto 99
20 print *, 'two'; goto 99
30 print *, 'three'
99 continue
end program t
