! vybe-test: fortran/sql_cics/sql_delete_05
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC SQL DELETE FROM T WHERE ID = 1 END-EXEC
end program p
