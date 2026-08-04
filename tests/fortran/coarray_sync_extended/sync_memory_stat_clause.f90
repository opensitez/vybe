! vybe-test: fortran/coarray_sync_extended/sync_memory_stat_clause
! origin: languages/fortran/tests/fortran/test_coarray_sync_extended.rs
program t
integer :: stat, x[*]
x = 0
sync memory (stat=stat)
x = 1
print *, x, stat
end program t
