! vybe-test: fortran/array_transforms/transpose_2x2_swap_off_diagonal
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(2,2)
integer :: b(2,2)
a(1,1)=9
a(1,2)=8
a(2,1)=7
a(2,2)=6
b=transpose(a)
if ((b(1,2)) /= 7) then
    print *, "FAIL: want [7] got [", b(1,2), "]"
    stop 1
end if
if ((b(2,1)) /= 8) then
    print *, "FAIL: want [8] got [", b(2,1), "]"
    stop 1
end if
if ((sum(b)) /= 30) then
    print *, "FAIL: want [30] got [", sum(b), "]"
    stop 1
end if
end program t
