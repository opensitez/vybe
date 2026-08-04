! vybe-test: fortran/where_merge_extended/where_multi_else_sign_three_way
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: v(3)=[-2,0,5]
integer :: s(3)
where (v<0)
s=-1
elsewhere (v==0)
s=0
elsewhere
s=1
end where
if ((s(1)) /= -1) then
    print *, "FAIL: want [-1] got [", s(1), "]"
    stop 1
end if
if ((s(2)) /= 0) then
    print *, "FAIL: want [0] got [", s(2), "]"
    stop 1
end if
if ((s(3)) /= 1) then
    print *, "FAIL: want [1] got [", s(3), "]"
    stop 1
end if
end program t
