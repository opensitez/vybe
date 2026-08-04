! vybe-test: fortran/sql_cics/sql_dynamic_20
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
character(len=40) :: stmt
EXEC SQL PREPARE S1 FROM :stmt END-EXEC
end program p
