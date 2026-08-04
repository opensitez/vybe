! vybe-test: fortran/array_transforms/merge_array_alternate_mask
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3)=[1,2,3]
integer :: b(3)=[4,5,6]
logical :: m(3)=[.true.,.false.,.true.]
integer :: c(3)
c=merge(a,b,m)
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
