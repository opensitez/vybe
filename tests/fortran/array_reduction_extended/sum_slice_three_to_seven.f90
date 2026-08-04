! vybe-test: fortran/array_reduction_extended/sum_slice_three_to_seven
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(9) = [(i * 2, i = 1, 9)]
if ((sum(a(3:7))) /= 50) then
    print *, "FAIL: want [50] got [", sum(a(3:7)), "]"
    stop 1
end if
end program t
