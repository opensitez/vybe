! vybe-test: fortran/enum_type_extended/enum_dtype_array_field
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
enum, bind(c)
enumerator :: A = 0, B = 1, C = 2
end enum
type :: Set
integer :: tags(3)
end type Set
type(Set) :: s
s%tags = [A, B, C]
if ((s%tags(B + 1)) /= 1) then
    print *, "FAIL: want [1] got [", s%tags(B + 1), "]"
    stop 1
end if
end program t
