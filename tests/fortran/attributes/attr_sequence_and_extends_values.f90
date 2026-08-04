! vybe-test: fortran/attributes/attr_sequence_and_extends_values
! origin: languages/fortran/tests/fortran/test_attributes.rs

program attr_sequence_and_extends_values
    type :: Base
        integer :: base_field = 5
    end type Base

    type, extends(Base) :: Child
        integer :: child_field = 7
    end type Child

    type(Child) :: c
    if ((c%base_field) /= 5) then
    print *, "FAIL: want [5] got [", c%base_field, "]"
    stop 1
end if
    if ((c%child_field) /= 7) then
    print *, "FAIL: want [7] got [", c%child_field, "]"
    stop 1
end if
    if ((c%base_field + c%child_field) /= 12) then
    print *, "FAIL: want [12] got [", c%base_field + c%child_field, "]"
    stop 1
end if
end program attr_sequence_and_extends_values
