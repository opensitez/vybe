! vybe-test: fortran/where_merge_extended/nested_where_even_odd_tiers
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[2,3,4,5]
integer :: b(4)=0
where (a>2)
where (mod(a,2)==0)
b=a*10
elsewhere
b=a
end where
end where
if ((b(1)) /= 0) then
    print *, "FAIL: want [0] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 3) then
    print *, "FAIL: want [3] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 40) then
    print *, "FAIL: want [40] got [", b(3), "]"
    stop 1
end if
end program t
