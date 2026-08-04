! vybe-test: fortran/sql_cics/cics_syncpoint_26
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC CICS SYNCPOINT
EXEC CICS SYNCPOINT ROLLBACK
EXEC CICS SYNCPOINT ROLLBACK(YES)
EXEC CICS SYNCPOINT KEEP
end program p
