! vybe-test: fortran/array_reduction_extended/sum_int_constant_vector
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(5) = [7, 7, 7, 7, 7]
if ((sum(a)) /= 35) then
    print *, "FAIL: want [35] got [", sum(a), "]"
    stop 1
end if
end program t
