! vybe-test: fortran/where_merge_extended/nested_where_2d_positive_inner
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: m(2,2)=reshape([1,-2,3,-4],[2,2])
integer :: r(2,2)=0
where (m>0)
where (m>2)
r=m*2
elsewhere
r=m
end where
end where
if ((r(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", r(1,1), "]"
    stop 1
end if
if ((r(2,1)) /= 0) then
    print *, "FAIL: want [0] got [", r(2,1), "]"
    stop 1
end if
if ((r(2,2)) /= 0) then
    print *, "FAIL: want [0] got [", r(2,2), "]"
    stop 1
end if
end program t
