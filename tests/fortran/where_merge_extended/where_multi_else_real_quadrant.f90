! vybe-test: fortran/where_merge_extended/where_multi_else_real_quadrant
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
real :: x(4)=[-3.0,2.0,-1.0,4.0]
integer :: q(4)
where (x<0.0)
q=1
elsewhere (x<3.0)
q=2
elsewhere
q=3
end where
if ((q(1)) /= 1) then
    print *, "FAIL: want [1] got [", q(1), "]"
    stop 1
end if
if ((q(2)) /= 2) then
    print *, "FAIL: want [2] got [", q(2), "]"
    stop 1
end if
if ((q(4)) /= 3) then
    print *, "FAIL: want [3] got [", q(4), "]"
    stop 1
end if
end program t
