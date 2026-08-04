! vybe-test: fortran/sql_cics/sql_insert_03
! origin: languages/fortran/tests/fortran/test_sql_cics.rs
program p
implicit none
integer :: id
id = 1
EXEC SQL INSERT INTO T(ID) VALUES(:id) END-EXEC
end program p
