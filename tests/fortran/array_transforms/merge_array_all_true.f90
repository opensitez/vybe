! vybe-test: fortran/array_transforms/merge_array_all_true
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3)=[8,8,8]
integer :: b(3)=[1,2,3]
logical :: m(3)=[.true.,.true.,.true.]
integer :: c(3)
c=merge(a,b,m)
if ((c(1)) /= 8) then
    print *, "FAIL: want [8] got [", c(1), "]"
    stop 1
end if
if ((c(3)) /= 8) then
    print *, "FAIL: want [8] got [", c(3), "]"
    stop 1
end if
if ((sum(c)) /= 24) then
    print *, "FAIL: want [24] got [", sum(c), "]"
    stop 1
end if
end program t
