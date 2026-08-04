! vybe-test: fortran/sql_cics/sql_cursor_close_21
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
integer :: rc
EXEC SQL CLOSE C1 INTO :rc END-EXEC
end program p
