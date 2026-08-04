! vybe-test: fortran/array_reduction_extended/sum_int_alternating_signs
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(6) = [5, -2, 8, -1, 3, -4]
if ((sum(a)) /= 9) then
    print *, "FAIL: want [9] got [", sum(a), "]"
    stop 1
end if
end program t
