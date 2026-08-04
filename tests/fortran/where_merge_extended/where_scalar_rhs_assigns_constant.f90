! vybe-test: fortran/where_merge_extended/where_scalar_rhs_assigns_constant
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[1,2,3,4]
where (a>2)
a=99
end where
if ((a(2)) /= 2) then
    print *, "FAIL: want [2] got [", a(2), "]"
    stop 1
end if
if ((a(3)) /= 99) then
    print *, "FAIL: want [99] got [", a(3), "]"
    stop 1
end if
if ((a(4)) /= 99) then
    print *, "FAIL: want [99] got [", a(4), "]"
    stop 1
end if
end program t
