! vybe-test: fortran/where_merge_extended/merge_array_from_comparison_mask
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[1,2,3,4]
integer :: b(4)=[10,20,30,40]
integer :: c(4)
c=merge(a,b,a>2)
if ((c(1)) /= 10) then
    print *, "FAIL: want [10] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= 20) then
    print *, "FAIL: want [20] got [", c(2), "]"
    stop 1
end if
if ((c(4)) /= 4) then
    print *, "FAIL: want [4] got [", c(4), "]"
    stop 1
end if
end program t
