! vybe-test: fortran/deallocate_statement/deallocate_statement_runtime_frees_allocated_scalar
! origin: languages/fortran/tests/fortran/test_deallocate_statement.rs
program t
integer, allocatable :: x
allocate(x)
x = 11
deallocate(x)
if (trim('deallocated') /= "deallocated") then
    print *, "FAIL: want [deallocated] got [", 'deallocated', "]"
    stop 1
end if
end program t
