! vybe-test: fortran/where_merge_extended/where_else_2d_hi_lo_split
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: m(2,2)=reshape([1,6,3,8],[2,2])
integer :: r(2,2)
where (m>5)
r=m*2
elsewhere
r=m
end where
if ((r(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", r(1,1), "]"
    stop 1
end if
if ((r(1,2)) /= 3) then
    print *, "FAIL: want [3] got [", r(1,2), "]"
    stop 1
end if
if ((r(2,2)) /= 16) then
    print *, "FAIL: want [16] got [", r(2,2), "]"
    stop 1
end if
end program t
