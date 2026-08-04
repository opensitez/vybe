! vybe-test: fortran/where_merge_extended/where_masked_abs_copy
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[-5,3,-2,7]
integer :: b(4)=0
where (a<0)
b=-a
end where
if ((b(1)) /= 5) then
    print *, "FAIL: want [5] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 0) then
    print *, "FAIL: want [0] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 2) then
    print *, "FAIL: want [2] got [", b(3), "]"
    stop 1
end if
end program t
