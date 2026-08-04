! vybe-test: fortran/where_merge_extended/merge_real_scalar_false_branch
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
real :: x
x=merge(3.5,7.5,.false.)
if (abs((x) - 7.5) > 1.0e-6) then
    print *, "FAIL: want [7.5] got [", x, "]"
    stop 1
end if
end program t
