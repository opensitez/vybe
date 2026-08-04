! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_computed_goto
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_computed_goto
integer n
n = 2
go to (10, 20, 30), n
10          print *, 10
go to 99
20          print *, 20
go to 99
30          print *, 30
+99          continue
end program statement_f77_legacy_compat_computed_goto
