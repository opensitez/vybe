! vybe-test: fortran/array_reduction_extended/sum_int_sparse_nonzeros
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(8) = [0, 0, 10, 0, 20, 0, 30, 0]
if ((sum(a)) /= 60) then
    print *, "FAIL: want [60] got [", sum(a), "]"
    stop 1
end if
end program t
