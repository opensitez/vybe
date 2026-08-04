! vybe-test: fortran/array_transforms/transpose_2x2_identity_like
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(2,2)
integer :: b(2,2)
a(1,1)=1
a(1,2)=0
a(2,1)=0
a(2,2)=1
b=transpose(a)
if ((b(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1,1), "]"
    stop 1
end if
if ((b(2,2)) /= 1) then
    print *, "FAIL: want [1] got [", b(2,2), "]"
    stop 1
end if
if ((sum(b)) /= 2) then
    print *, "FAIL: want [2] got [", sum(b), "]"
    stop 1
end if
end program t
