! vybe-test: fortran/sql_cics/sql_cursor_decl_08
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC SQL DECLARE C1 CURSOR FOR SELECT ID FROM T END-EXEC
end program p
