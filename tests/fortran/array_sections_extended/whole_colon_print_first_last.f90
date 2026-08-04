! vybe-test: fortran/array_sections_extended/whole_colon_print_first_last
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(6) = [(i * 3, i = 1, 6)]
if ((a(:)(1)) /= 3) then
    print *, "FAIL: want [3] got [", a(:)(1), "]"
    stop 1
end if
if ((a(:)(6)) /= 18) then
    print *, "FAIL: want [18] got [", a(:)(6), "]"
    stop 1
end if
end program t
