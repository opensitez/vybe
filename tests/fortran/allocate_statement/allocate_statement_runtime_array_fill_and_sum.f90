! vybe-test: fortran/allocate_statement/allocate_statement_runtime_array_fill_and_sum
! origin: languages/fortran/tests/fortran/test_allocate_statement.rs
program t
integer, allocatable :: a(:)
allocate(a(3))
a = [10, 20, 30]
if ((sum(a)) /= 60) then
    print *, "FAIL: want [60] got [", sum(a), "]"
    stop 1
end if
end program t
