! vybe-test: fortran/where_merge_extended/where_multi_else_temp_four_ranges
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
real :: t(4)=[-10.0,-1.0,5.0,50.0]
integer :: c(4)
where (t<0.0)
c=0
elsewhere (t<10.0)
c=1
elsewhere (t<40.0)
c=2
elsewhere
c=3
end where
if ((c(1)) /= 0) then
    print *, "FAIL: want [0] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= 1) then
    print *, "FAIL: want [1] got [", c(2), "]"
    stop 1
end if
if ((c(3)) /= 1) then
    print *, "FAIL: want [1] got [", c(3), "]"
    stop 1
end if
if ((c(4)) /= 3) then
    print *, "FAIL: want [3] got [", c(4), "]"
    stop 1
end if
end program t
