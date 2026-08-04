! vybe-test: fortran/where_merge_extended/merge_mask_from_scalar_flag
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
logical :: flag
integer :: x
flag = .true.
x = merge(10, 20, flag)
if ((x) /= 10) then
    print *, "FAIL: want [10] got [", x, "]"
    stop 1
end if
flag = .false.
x = merge(10, 20, flag)
if ((x) /= 20) then
    print *, "FAIL: want [20] got [", x, "]"
    stop 1
end if
end program t
