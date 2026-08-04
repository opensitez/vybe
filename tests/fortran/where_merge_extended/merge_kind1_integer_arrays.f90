! vybe-test: fortran/where_merge_extended/merge_kind1_integer_arrays
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer(kind=1) :: a(3)=[1_1, 2_1, 3_1]
integer(kind=1) :: b(3)=[4_1, 5_1, 6_1]
logical :: m(3)=[.true., .false., .true.]
integer(kind=1) :: c(3)
c = merge(a, b, m)
if ((c(1)) /= 1) then
    print *, "FAIL: want [1] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= 5) then
    print *, "FAIL: want [5] got [", c(2), "]"
    stop 1
end if
if ((c(3)) /= 3) then
    print *, "FAIL: want [3] got [", c(3), "]"
    stop 1
end if
end program t
