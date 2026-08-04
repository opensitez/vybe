! vybe-test: fortran/sql_cics/sql_cursor_fetch_10
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
integer :: id
EXEC SQL FETCH C1 INTO :id END-EXEC
end program p
