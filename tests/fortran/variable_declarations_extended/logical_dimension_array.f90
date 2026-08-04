! vybe-test: fortran/variable_declarations_extended/logical_dimension_array
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
logical, dimension(3) :: flags
flags(2) = .true.
if ((flags(2)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", flags(2), "]"
    stop 1
end if
end program t
