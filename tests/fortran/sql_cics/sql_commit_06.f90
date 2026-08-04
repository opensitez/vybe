! vybe-test: fortran/sql_cics/sql_commit_06
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC SQL COMMIT END-EXEC
end program p
