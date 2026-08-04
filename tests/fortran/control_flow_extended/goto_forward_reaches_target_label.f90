! vybe-test: fortran/control_flow_extended/goto_forward_reaches_target_label
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
goto 20
10 print *, 'skip'
goto 30
20 print *, 'landed'
30 continue
end program t
