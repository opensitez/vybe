! vybe-test: fortran/where_merge_extended/where_else_real_clamp_negatives
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
real :: x(4)=[-1.0,2.0,-3.0,4.0]
real :: y(4)
where (x>=0.0)
y=x
elsewhere
y=0.0
end where
if ((y(1)) /= 0) then
    print *, "FAIL: want [0] got [", y(1), "]"
    stop 1
end if
if ((y(2)) /= 2) then
    print *, "FAIL: want [2] got [", y(2), "]"
    stop 1
end if
if ((y(3)) /= 0) then
    print *, "FAIL: want [0] got [", y(3), "]"
    stop 1
end if
end program t
