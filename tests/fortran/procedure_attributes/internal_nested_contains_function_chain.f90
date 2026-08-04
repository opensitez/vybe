! vybe-test: fortran/procedure_attributes/internal_nested_contains_function_chain
! origin: languages/fortran/tests/fortran/test_procedure_attributes.rs
program t
call outer_driver()
contains
subroutine outer_driver()
call middle_layer(7)
contains
subroutine middle_layer(v)
integer, intent(in) :: v
if ((inner_fn(v)) /= 10) then
    print *, "FAIL: want [10] got [", inner_fn(v), "]"
    stop 1
end if
contains
function inner_fn(n) result(r)
integer, intent(in) :: n
integer :: r
r = n + 3
end function inner_fn
end subroutine middle_layer
end subroutine outer_driver
end program t
