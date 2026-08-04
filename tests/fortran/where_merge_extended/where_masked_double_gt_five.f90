! vybe-test: fortran/where_merge_extended/where_masked_double_gt_five
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(5)=[2,6,3,8,1]
where (a>5)
a=a*2
end where
if ((a(2)) /= 12) then
    print *, "FAIL: want [12] got [", a(2), "]"
    stop 1
end if
if ((a(3)) /= 3) then
    print *, "FAIL: want [3] got [", a(3), "]"
    stop 1
end if
if ((a(4)) /= 16) then
    print *, "FAIL: want [16] got [", a(4), "]"
    stop 1
end if
end program t
