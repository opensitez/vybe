! vybe-test: fortran/array_transforms/transpose_2x2_antidiagonal
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(2,2)
integer :: b(2,2)
a(1,1)=0
a(1,2)=5
a(2,1)=7
a(2,2)=0
b=transpose(a)
if ((b(1,2)) /= 7) then
    print *, "FAIL: want [7] got [", b(1,2), "]"
    stop 1
end if
if ((b(2,1)) /= 5) then
    print *, "FAIL: want [5] got [", b(2,1), "]"
    stop 1
end if
if ((sum(b)) /= 12) then
    print *, "FAIL: want [12] got [", sum(b), "]"
    stop 1
end if
end program t
