! vybe-test: fortran/array_transforms/merge_array_single_true_middle
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3)=[10,20,30]
integer :: b(3)=[1,2,3]
logical :: m(3)=[.false.,.true.,.false.]
integer :: c(3)
c=merge(a,b,m)
if ((c(1)) /= 1) then
    print *, "FAIL: want [1] got [", c(1), "]"
    stop 1
end if
if ((c(2)) /= 20) then
    print *, "FAIL: want [20] got [", c(2), "]"
    stop 1
end if
if ((c(3)) /= 3) then
    print *, "FAIL: want [3] got [", c(3), "]"
    stop 1
end if
end program t
