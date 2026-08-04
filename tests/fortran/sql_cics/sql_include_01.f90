! vybe-test: fortran/sql_cics/sql_include_01
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
EXEC SQL INCLUDE SQLCA END-EXEC
print *, 1
end program p
