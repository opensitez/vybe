! vybe-test: fortran/variable_declarations_extended/init_logical_from_comparison
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
logical :: gt = (5 > 3)
if ((gt) .neqv. .true.) then
    print *, "FAIL: want [true] got [", gt, "]"
    stop 1
end if
end program t
