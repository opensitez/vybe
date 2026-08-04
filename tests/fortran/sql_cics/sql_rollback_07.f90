! vybe-test: fortran/sql_cics/sql_rollback_07
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC SQL ROLLBACK END-EXEC
end program p
