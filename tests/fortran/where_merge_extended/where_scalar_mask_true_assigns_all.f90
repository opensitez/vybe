! vybe-test: fortran/where_merge_extended/where_scalar_mask_true_assigns_all
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(3)=[1,2,3]
where (.true.)
a=9
end where
if ((a(1)) /= 9) then
    print *, "FAIL: want [9] got [", a(1), "]"
    stop 1
end if
if ((a(3)) /= 9) then
    print *, "FAIL: want [9] got [", a(3), "]"
    stop 1
end if
if ((sum(a)) /= 27) then
    print *, "FAIL: want [27] got [", sum(a), "]"
    stop 1
end if
end program t
