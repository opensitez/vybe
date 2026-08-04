! vybe-test: fortran/where_merge_extended/merge_mask_from_variable
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
logical :: flag=.false.
integer :: x
x=merge(100,200,flag)
if ((x) /= 200) then
    print *, "FAIL: want [200] got [", x, "]"
    stop 1
end if
end program t
