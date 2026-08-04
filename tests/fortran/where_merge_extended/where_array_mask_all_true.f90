! vybe-test: fortran/where_merge_extended/where_array_mask_all_true
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(3)=[1,2,3]
integer :: b(3)=0
where (a>0)
b=a*2
end where
if ((b(2)) /= 4) then
    print *, "FAIL: want [4] got [", b(2), "]"
    stop 1
end if
if ((sum(b)) /= 12) then
    print *, "FAIL: want [12] got [", sum(b), "]"
    stop 1
end if
end program t
