! vybe-test: fortran/where_merge_extended/nested_where_with_merge_mask
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(4)=[1,2,3,4]
integer :: b(4)=[4,3,2,1]
logical :: m(4)
m = merge((/ .true., .false., .false., .true. /), (/ .false., .true., .true., .false. /), a > 2)
where (m)
b = a * 10
elsewhere
b = b - a
end where
if ((b(1)) /= 3) then
    print *, "FAIL: want [3] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 20) then
    print *, "FAIL: want [20] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= -1) then
    print *, "FAIL: want [-1] got [", b(3), "]"
    stop 1
end if
if ((b(4)) /= 40) then
    print *, "FAIL: want [40] got [", b(4), "]"
    stop 1
end if
end program t
