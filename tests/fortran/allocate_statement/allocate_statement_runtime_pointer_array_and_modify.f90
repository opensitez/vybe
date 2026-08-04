! vybe-test: fortran/allocate_statement/allocate_statement_runtime_pointer_array_and_modify
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program t
integer, pointer :: p(:)
allocate(p(2))
p(1) = 7
p(2) = 9
if ((p(2)) /= 9) then
    print *, "FAIL: want [9] got [", p(2), "]"
    stop 1
end if
deallocate(p)
if (trim('done') /= "done") then
    print *, "FAIL: want [done] got [", 'done', "]"
    stop 1
end if
end program t
