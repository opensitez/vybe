! vybe-test: fortran/procedure_attributes/elemental_custom_abs_diff_on_arrays
! origin: languages/fortran/tests/fortran/test_procedure_attributes.rs
program t
integer :: a(3), b(3), c(3)
a = [5, 1, 9]
b = [2, 4, 3]
c = abs_diff(a, b)
if ((c(1)) /= 3) then
    print *, "FAIL: want [3] got [", c(1), "]"
    stop 1
end if
if ((sum(c)) /= 10) then
    print *, "FAIL: want [10] got [", sum(c), "]"
    stop 1
end if
contains
elemental function abs_diff(x, y) result(d)
integer, intent(in) :: x, y
integer :: d
d = abs(x - y)
end function abs_diff
end program t
