! vybe-test: fortran/where_merge_extended/where_masked_assign_from_other_array
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: src(4)=[10,20,30,40]
integer :: dst(4)=[0,0,0,0]
integer :: mask(4)=[1,0,1,0]
where (mask==1)
dst=src
end where
if ((dst(1)) /= 10) then
    print *, "FAIL: want [10] got [", dst(1), "]"
    stop 1
end if
if ((dst(2)) /= 0) then
    print *, "FAIL: want [0] got [", dst(2), "]"
    stop 1
end if
if ((dst(3)) /= 30) then
    print *, "FAIL: want [30] got [", dst(3), "]"
    stop 1
end if
end program t
