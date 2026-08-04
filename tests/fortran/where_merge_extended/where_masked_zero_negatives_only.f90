! vybe-test: fortran/where_merge_extended/where_masked_zero_negatives_only
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(5)=[3,-1,7,-4,2]
where (a<0)
a=0
end where
if ((a(2)) /= 0) then
    print *, "FAIL: want [0] got [", a(2), "]"
    stop 1
end if
if ((a(3)) /= 7) then
    print *, "FAIL: want [7] got [", a(3), "]"
    stop 1
end if
if ((sum(a)) /= 12) then
    print *, "FAIL: want [12] got [", sum(a), "]"
    stop 1
end if
end program t
