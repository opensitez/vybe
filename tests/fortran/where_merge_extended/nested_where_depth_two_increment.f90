! vybe-test: fortran/where_merge_extended/nested_where_depth_two_increment
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[1,6,11,16]
integer :: b(4)=0
where (a>3)
where (a>10)
b=b+2
elsewhere
b=b+1
end where
end where
if ((b(1)) /= 0) then
    print *, "FAIL: want [0] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 1) then
    print *, "FAIL: want [1] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 2) then
    print *, "FAIL: want [2] got [", b(3), "]"
    stop 1
end if
end program t
