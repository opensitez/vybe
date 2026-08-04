! vybe-test: fortran/where_merge_extended/where_else_logical_int_encoding
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[-1,0,2,5]
integer :: b(4)
where (a>0)
b=1
elsewhere
b=0
end where
if ((b(1)) /= 0) then
    print *, "FAIL: want [0] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 0) then
    print *, "FAIL: want [0] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 1) then
    print *, "FAIL: want [1] got [", b(3), "]"
    stop 1
end if
end program t
