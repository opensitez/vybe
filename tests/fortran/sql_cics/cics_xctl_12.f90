! vybe-test: fortran/sql_cics/cics_xctl_12
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC CICS XCTL PROGRAM('SUB2') END-EXEC
end program p
