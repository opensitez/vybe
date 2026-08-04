! vybe-test: fortran/control_flow_extended/go_to_spelling_variant_reaches_label
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
go to 20
10 print *, 'miss'
20 print *, 'hit'
end program t
