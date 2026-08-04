! vybe-test: fortran/associate_construct_extended/associate_array_section_sum
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a(6) = [1, 2, 3, 4, 5, 6]
associate (slice => a(2:4))
if ((sum(slice)) /= 9) then
    print *, "FAIL: want [9] got [", sum(slice), "]"
    stop 1
end if
end associate
end program t
