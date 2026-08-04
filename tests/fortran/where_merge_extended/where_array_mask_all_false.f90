! vybe-test: fortran/where_merge_extended/where_array_mask_all_false
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(3)=[5,6,7]
integer :: b(3)=0
where (a>10)
b=1
end where
if ((b(1)) /= 0) then
    print *, "FAIL: want [0] got [", b(1), "]"
    stop 1
end if
if ((sum(b)) /= 0) then
    print *, "FAIL: want [0] got [", sum(b), "]"
    stop 1
end if
end program t
