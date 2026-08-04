! vybe-test: fortran/implied_do_forms/ac_zero_to_five_even_stride
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
integer :: a(3) = [(i, i = 0, 5, 2)]
if ((sum(a)) /= 6) then
    print *, "FAIL: want [6] got [", sum(a), "]"
    stop 1
end if
end program t
