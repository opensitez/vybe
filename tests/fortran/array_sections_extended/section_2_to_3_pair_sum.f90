! vybe-test: fortran/array_sections_extended/section_2_to_3_pair_sum
! origin: languages/fortran/tests/fortran/test_array_sections_extended.rs
program t
integer :: a(5) = [10, 20, 30, 40, 50]
if ((sum(a(2:3))) /= 50) then
    print *, "FAIL: want [50] got [", sum(a(2:3)), "]"
    stop 1
end if
end program t
