! vybe-test: fortran/where_merge_extended/merge_in_where_body
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(3)=[-1,2,-3]
integer :: b(3)
where (a<0)
b=merge(0,a,.true.)
elsewhere
b=a
end where
if ((b(1)) /= 0) then
    print *, "FAIL: want [0] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 2) then
    print *, "FAIL: want [2] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 0) then
    print *, "FAIL: want [0] got [", b(3), "]"
    stop 1
end if
end program t
