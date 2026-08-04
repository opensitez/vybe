! vybe-test: fortran/associate_construct_extended/associate_dtype_field_modify
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
type :: Counter
integer :: n = 0
end type Counter
type(Counter) :: c
associate (count => c%n)
count = 15
end associate
if ((c%n) /= 15) then
    print *, "FAIL: want [15] got [", c%n, "]"
    stop 1
end if
end program t
