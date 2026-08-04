! vybe-test: fortran/sql_cics/cics_writeq_ts_17
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC CICS WRITEQ TS QUEUE('Q1') END-EXEC
end program p
