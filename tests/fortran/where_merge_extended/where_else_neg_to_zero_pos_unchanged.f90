! vybe-test: fortran/where_merge_extended/where_else_neg_to_zero_pos_unchanged
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: v(5)=[5,-2,8,-1,3]
where (v<0)
v=0
elsewhere
v=v
end where
if ((v(2)) /= 0) then
    print *, "FAIL: want [0] got [", v(2), "]"
    stop 1
end if
if ((v(3)) /= 8) then
    print *, "FAIL: want [8] got [", v(3), "]"
    stop 1
end if
if ((sum(v)) /= 16) then
    print *, "FAIL: want [16] got [", sum(v), "]"
    stop 1
end if
end program t
