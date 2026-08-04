! vybe-test: fortran/derived_type_oop_extended/extends_shadow_child_integer_wins
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs
program t
type :: Base
integer :: val = 1
end type Base
type, extends(Base) :: Derived
integer :: val = 99
end type Derived
type(Derived) :: d
if ((d%val) /= 99) then
    print *, "FAIL: want [99] got [", d%val, "]"
    stop 1
end if
end program t
