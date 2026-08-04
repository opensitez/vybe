! vybe-test: fortran/array_reduction_extended/sum_int_triangular_ten
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(4) = [1, 3, 6, 10]
if ((sum(a)) /= 20) then
    print *, "FAIL: want [20] got [", sum(a), "]"
    stop 1
end if
end program t
