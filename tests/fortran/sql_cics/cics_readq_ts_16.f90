! vybe-test: fortran/sql_cics/cics_readq_ts_16
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC CICS READQ TS QUEUE('Q1') END-EXEC
end program p
