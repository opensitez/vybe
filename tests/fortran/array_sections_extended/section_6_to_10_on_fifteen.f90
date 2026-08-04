! vybe-test: fortran/array_sections_extended/section_6_to_10_on_fifteen
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(15) = [(i, i = 1, 15)]
if ((a(6:10)(1)) /= 6) then
    print *, "FAIL: want [6] got [", a(6:10)(1), "]"
    stop 1
end if
if ((a(6:10)(5)) /= 10) then
    print *, "FAIL: want [10] got [", a(6:10)(5), "]"
    stop 1
end if
if ((sum(a(6:10))) /= 40) then
    print *, "FAIL: want [40] got [", sum(a(6:10)), "]"
    stop 1
end if
end program t
