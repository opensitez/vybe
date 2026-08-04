! vybe-test: fortran/array_transforms/transpose_2x3_row_matrix_sum
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3,2)
integer :: b(2,3)
a(1,1)=1
a(1,2)=2
a(2,1)=3
a(2,2)=4
a(3,1)=5
a(3,2)=6
b=transpose(a)
if ((b(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1,1), "]"
    stop 1
end if
if ((b(2,3)) /= 6) then
    print *, "FAIL: want [6] got [", b(2,3), "]"
    stop 1
end if
if ((sum(b)) /= 21) then
    print *, "FAIL: want [21] got [", sum(b), "]"
    stop 1
end if
end program t
