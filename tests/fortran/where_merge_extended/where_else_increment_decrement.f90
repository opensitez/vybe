! vybe-test: fortran/where_merge_extended/where_else_increment_decrement
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[1,4,7,10]
where (mod(a,2)==0)
a=a+1
elsewhere
a=a-1
end where
if ((a(1)) /= 0) then
    print *, "FAIL: want [0] got [", a(1), "]"
    stop 1
end if
if ((a(2)) /= 5) then
    print *, "FAIL: want [5] got [", a(2), "]"
    stop 1
end if
if ((a(3)) /= 6) then
    print *, "FAIL: want [6] got [", a(3), "]"
    stop 1
end if
end program t
