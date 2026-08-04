! vybe-test: fortran/associate_construct_extended/associate_array_section_first
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a(5) = [10, 20, 30, 40, 50]
associate (slice => a(1:2))
if ((slice(2)) /= 20) then
    print *, "FAIL: want [20] got [", slice(2), "]"
    stop 1
end if
end associate
end program t
