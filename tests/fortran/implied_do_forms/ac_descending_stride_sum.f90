! vybe-test: fortran/implied_do_forms/ac_descending_stride_sum
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
integer :: a(5) = [(i, i = 5, 1, -1)]
if ((sum(a)) /= 15) then
    print *, "FAIL: want [15] got [", sum(a), "]"
    stop 1
end if
end program t
