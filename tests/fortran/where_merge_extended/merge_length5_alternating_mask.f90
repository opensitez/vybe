! vybe-test: fortran/where_merge_extended/merge_length5_alternating_mask
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: a(5)=[1,1,1,1,1]
integer :: b(5)=[2,2,2,2,2]
logical :: m(5)=[.true.,.false.,.true.,.false.,.true.]
integer :: c(5)
c=merge(a,b,m)
if ((c(1)) /= 1) then
    print *, "FAIL: want [1] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= 2) then
    print *, "FAIL: want [2] got [", c(2), "]"
    stop 1
end if
if ((c(5)) /= 1) then
    print *, "FAIL: want [1] got [", c(5), "]"
    stop 1
end if
if ((sum(c)) /= 7) then
    print *, "FAIL: want [7] got [", sum(c), "]"
    stop 1
end if
end program t
