! vybe-test: fortran/where_merge_extended/where_masked_self_increment
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[1,2,3,4]
where (a>=3)
a=a+10
end where
if ((a(3)) /= 13) then
    print *, "FAIL: want [13] got [", a(3), "]"
    stop 1
end if
if ((a(4)) /= 14) then
    print *, "FAIL: want [14] got [", a(4), "]"
    stop 1
end if
if ((sum(a)) /= 30) then
    print *, "FAIL: want [30] got [", sum(a), "]"
    stop 1
end if
end program t
