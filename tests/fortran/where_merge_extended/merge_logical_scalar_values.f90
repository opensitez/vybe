! vybe-test: fortran/where_merge_extended/merge_logical_scalar_values
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
logical :: x
x=merge(.true.,.false.,.true.)
if ((x) .neqv. .true.) then
    print *, "FAIL: want [true] got [", x, "]"
    stop 1
end if
end program t
