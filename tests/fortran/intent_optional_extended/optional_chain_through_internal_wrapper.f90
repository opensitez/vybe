! vybe-test: fortran/intent_optional_extended/optional_chain_through_internal_wrapper
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if ((wrap_add(2)) /= 2) then
    print *, "FAIL: want [2] got [", wrap_add(2), "]"
    stop 1
end if
if ((wrap_add(2, 8)) /= 10) then
    print *, "FAIL: want [10] got [", wrap_add(2, 8), "]"
    stop 1
end if
contains
function inner_add(a, b) result(r)
integer, intent(in) :: a
integer, intent(in), optional :: b
integer :: r
r = a
if (present(b)) r = r + b
end function inner_add
integer function wrap_add(x, y)
integer, intent(in) :: x
integer, intent(in), optional :: y
if (present(y)) then
wrap_add = inner_add(x, y)
else
wrap_add = inner_add(x)
end if
end function wrap_add
end program t
