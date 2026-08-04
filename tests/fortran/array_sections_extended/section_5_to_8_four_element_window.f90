! vybe-test: fortran/array_sections_extended/section_5_to_8_four_element_window
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(10) = [(i * 10, i = 1, 10)]
if ((a(5:8)(1)) /= 50) then
    print *, "FAIL: want [50] got [", a(5:8)(1), "]"
    stop 1
end if
if ((a(5:8)(4)) /= 80) then
    print *, "FAIL: want [80] got [", a(5:8)(4), "]"
    stop 1
end if
if ((sum(a(5:8))) /= 260) then
    print *, "FAIL: want [260] got [", sum(a(5:8)), "]"
    stop 1
end if
end program t
