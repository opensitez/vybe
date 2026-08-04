! vybe-test: fortran/where_merge_extended/merge_real_scalar_true_branch
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
real :: x
x=merge(3.5,7.5,.true.)
if (abs((x) - 3.5) > 1.0e-6) then
    print *, "FAIL: want [3.5] got [", x, "]"
    stop 1
end if
end program t
