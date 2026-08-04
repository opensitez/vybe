! vybe-test: fortran/array_transforms/transpose_3x2_sequential_fill
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(2,3)
integer :: b(3,2)
a(1,1)=10
a(1,2)=20
a(1,3)=30
a(2,1)=40
a(2,2)=50
a(2,3)=60
b=transpose(a)
if ((b(1,1)) /= 10) then
    print *, "FAIL: want [10] got [", b(1,1), "]"
    stop 1
end if
if ((b(3,2)) /= 60) then
    print *, "FAIL: want [60] got [", b(3,2), "]"
    stop 1
end if
if ((sum(b)) /= 210) then
    print *, "FAIL: want [210] got [", sum(b), "]"
    stop 1
end if
end program t
