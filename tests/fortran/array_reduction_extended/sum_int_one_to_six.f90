! vybe-test: fortran/array_reduction_extended/sum_int_one_to_six
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(6) = [(i, i = 1, 6)]
if ((sum(a)) /= 21) then
    print *, "FAIL: want [21] got [", sum(a), "]"
    stop 1
end if
end program t
