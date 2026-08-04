! vybe-test: fortran/array_sections_extended/size_section_2_to_5_and_stride
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(10) = [(i, i = 1, 10)]
if ((size(a(2:5))) /= 4) then
    print *, "FAIL: want [4] got [", size(a(2:5)), "]"
    stop 1
end if
if ((size(a(1:10:2))) /= 5) then
    print *, "FAIL: want [5] got [", size(a(1:10:2)), "]"
    stop 1
end if
end program t
