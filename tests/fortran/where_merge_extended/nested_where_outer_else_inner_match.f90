! vybe-test: fortran/where_merge_extended/nested_where_outer_else_inner_match
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(3)=[10,3,20]
integer :: b(3)=0
where (a>5)
where (a>15)
b=2
elsewhere
b=1
end where
elsewhere
b=-1
end where
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= -1) then
    print *, "FAIL: want [-1] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 2) then
    print *, "FAIL: want [2] got [", b(3), "]"
    stop 1
end if
end program t
