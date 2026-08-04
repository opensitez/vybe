! vybe-test: fortran/associate_construct_extended/associate_dtype_integer_field
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
type :: Pair
integer :: x, y
end type Pair
type(Pair) :: p
p%x = 12
p%y = 30
associate (first => p%x)
if ((first) /= 12) then
    print *, "FAIL: want [12] got [", first, "]"
    stop 1
end if
end associate
end program t
