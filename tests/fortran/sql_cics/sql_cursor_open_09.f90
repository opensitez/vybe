! vybe-test: fortran/sql_cics/sql_cursor_open_09
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC SQL OPEN C1 END-EXEC
end program p
