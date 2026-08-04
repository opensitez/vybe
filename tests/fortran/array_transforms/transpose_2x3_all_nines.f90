! vybe-test: fortran/array_transforms/transpose_2x3_all_nines
! origin: languages/fortran/tests/fortran/test_array_transforms.rs
program t
integer :: a(3,2)
integer :: b(2,3)
a(1,1)=9
a(1,2)=9
a(2,1)=9
a(2,2)=9
a(3,1)=9
a(3,2)=9
b=transpose(a)
if ((b(2,2)) /= 9) then
    print *, "FAIL: want [9] got [", b(2,2), "]"
    stop 1
end if
if ((sum(b)) /= 54) then
    print *, "FAIL: want [54] got [", sum(b), "]"
    stop 1
end if
end program t
