! vybe-test: fortran/where_merge_extended/where_else_sum_of_branches
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[1,6,2,9]
integer :: b(4)
where (a>5)
b=a*2
elsewhere
b=a+1
end where
if ((b(1)) /= 2) then
    print *, "FAIL: want [2] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 12) then
    print *, "FAIL: want [12] got [", b(2), "]"
    stop 1
end if
if ((sum(b)) /= 35) then
    print *, "FAIL: want [35] got [", sum(b), "]"
    stop 1
end if
end program t
