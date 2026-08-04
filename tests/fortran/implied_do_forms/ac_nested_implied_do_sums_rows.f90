! vybe-test: fortran/implied_do_forms/ac_nested_implied_do_sums_rows
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
integer :: a(6) = [((i*10 + j, j = 1, 2), i = 1, 3)]
if ((sum(a)) /= 129) then
    print *, "FAIL: want [129] got [", sum(a), "]"
    stop 1
end if
end program t
