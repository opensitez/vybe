! vybe-test: fortran/deallocate_statement/deallocate_statement_runtime_pointer_is_releaseable
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program t
integer, pointer :: p(:)
allocate(p(3))
p = [1, 2, 3]
deallocate(p)
if (trim('done') /= "done") then
    print *, "FAIL: want [done] got [", 'done', "]"
    stop 1
end if
end program t
