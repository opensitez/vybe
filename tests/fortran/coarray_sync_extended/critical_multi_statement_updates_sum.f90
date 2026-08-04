! vybe-test: fortran/coarray_sync_extended/critical_multi_statement_updates_sum
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
integer :: a = 0, b = 0
critical
a = a + 2
b = b + a
end critical
print *, a, b
end program t
