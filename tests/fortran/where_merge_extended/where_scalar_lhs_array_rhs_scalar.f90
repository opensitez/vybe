! vybe-test: fortran/where_merge_extended/where_scalar_lhs_array_rhs_scalar
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(5)=[0,0,0,0,0]
integer :: b(5)=[1,2,3,4,5]
where (b>3)
a=b
end where
if ((a(4)) /= 4) then
    print *, "FAIL: want [4] got [", a(4), "]"
    stop 1
end if
if ((a(5)) /= 5) then
    print *, "FAIL: want [5] got [", a(5), "]"
    stop 1
end if
if ((sum(a)) /= 9) then
    print *, "FAIL: want [9] got [", sum(a), "]"
    stop 1
end if
end program t
