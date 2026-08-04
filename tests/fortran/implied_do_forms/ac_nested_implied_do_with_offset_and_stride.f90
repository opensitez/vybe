! vybe-test: fortran/implied_do_forms/ac_nested_implied_do_with_offset_and_stride
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
integer :: a(4) = [((i + j, j = 0, 2, 2), i = 1, 2)]
if ((sum(a)) /= 10) then
    print *, "FAIL: want [10] got [", sum(a), "]"
    stop 1
end if
end program t
