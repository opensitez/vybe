! vybe-test: fortran/where_merge_extended/where_scalar_mask_false_leaves_unchanged
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(3)=[4,5,6]
where (.false.)
a=0
end where
if ((a(1)) /= 4) then
    print *, "FAIL: want [4] got [", a(1), "]"
    stop 1
end if
if ((sum(a)) /= 15) then
    print *, "FAIL: want [15] got [", sum(a), "]"
    stop 1
end if
end program t
