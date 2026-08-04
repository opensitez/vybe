! vybe-test: fortran/where_merge_extended/where_masked_square_evens
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(6)=[1,2,3,4,5,6]
where (mod(a,2)==0)
a=a*a
end where
if ((a(2)) /= 4) then
    print *, "FAIL: want [4] got [", a(2), "]"
    stop 1
end if
if ((a(3)) /= 3) then
    print *, "FAIL: want [3] got [", a(3), "]"
    stop 1
end if
if ((a(4)) /= 16) then
    print *, "FAIL: want [16] got [", a(4), "]"
    stop 1
end if
end program t
