! vybe-test: fortran/where_merge_extended/merge_array_2d_mask
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(2,2)=reshape([1,2,3,4],[2,2])
integer :: b(2,2)=reshape([9,8,7,6],[2,2])
integer :: c(2,2)
c=merge(a,b,a<b)
if ((c(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", c(1,1), "]"
    stop 1
end if
if ((c(1,2)) /= 8) then
    print *, "FAIL: want [8] got [", c(1,2), "]"
    stop 1
end if
if ((c(2,1)) /= 3) then
    print *, "FAIL: want [3] got [", c(2,1), "]"
    stop 1
end if
end program t
