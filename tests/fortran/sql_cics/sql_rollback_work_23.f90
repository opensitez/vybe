! vybe-test: fortran/sql_cics/sql_rollback_work_23
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC SQL ROLLBACK WORK END-EXEC
end program p
