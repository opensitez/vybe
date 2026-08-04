! vybe-test: fortran/implied_do_forms/print_list_directed_from_implied_do
! origin: languages/fortran/tests/fortran/test_implied_do_forms.rs
program t
if (trim([(i, i = 1, 5)]) /= "1,2,3,4,5") then
    print *, "FAIL: want [1,2,3,4,5] got [", [(i, i = 1, 5)], "]"
    stop 1
end if
end program t
