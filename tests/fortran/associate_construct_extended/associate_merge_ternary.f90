! vybe-test: fortran/associate_construct_extended/associate_merge_ternary
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a = 10, b = 20
logical :: pick = .true.
associate (chosen => merge(a, b, pick))
if ((chosen) /= 10) then
    print *, "FAIL: want [10] got [", chosen, "]"
    stop 1
end if
end associate
end program t
