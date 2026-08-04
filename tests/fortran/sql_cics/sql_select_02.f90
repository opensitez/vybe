! vybe-test: fortran/sql_cics/sql_select_02
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
integer :: id
EXEC SQL SELECT 1 INTO :id END-EXEC
print *, id
end program p
