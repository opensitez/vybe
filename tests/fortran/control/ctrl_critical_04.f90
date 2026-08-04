! vybe-test: fortran/control/ctrl_critical_04
! origin: languages/fortran/tests/fortran/test_control.rs
program p
critical
 print *,1
end critical
end program p
