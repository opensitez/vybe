! vybe-test: fortran/where_merge_extended/where_else_int_multiply_branch
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[2,8,3,9]
integer :: b(4)
where (a>5)
b=a*10
elsewhere
b=a
end where
if ((b(1)) /= 2) then
    print *, "FAIL: want [2] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 80) then
    print *, "FAIL: want [80] got [", b(2), "]"
    stop 1
end if
if ((b(4)) /= 90) then
    print *, "FAIL: want [90] got [", b(4), "]"
    stop 1
end if
end program t
