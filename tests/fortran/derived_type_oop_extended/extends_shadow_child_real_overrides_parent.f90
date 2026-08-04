! vybe-test: fortran/derived_type_oop_extended/extends_shadow_child_real_overrides_parent
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs
program t
type :: Base
real :: metric = 1.0
end type Base
type, extends(Base) :: Derived
real :: metric = 3.5
end type Derived
type(Derived) :: d
if ((int(d%metric)) /= 3) then
    print *, "FAIL: want [3] got [", int(d%metric), "]"
    stop 1
end if
end program t
