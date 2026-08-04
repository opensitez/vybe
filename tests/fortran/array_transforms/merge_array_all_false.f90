! vybe-test: fortran/array_transforms/merge_array_all_false
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3)=[8,8,8]
integer :: b(3)=[1,2,3]
logical :: m(3)=[.false.,.false.,.false.]
integer :: c(3)
c=merge(a,b,m)
if ((c(1)) /= 1) then
    print *, "FAIL: want [1] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= 2) then
    print *, "FAIL: want [2] got [", c(2), "]"
    stop 1
end if
if ((c(3)) /= 3) then
    print *, "FAIL: want [3] got [", c(3), "]"
    stop 1
end if
end program t
