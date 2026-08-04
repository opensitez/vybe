! vybe-test: fortran/variable_declarations_extended/dimension_with_kind_attribute
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer(kind=4), dimension(3) :: arr
arr(2) = 77
if ((arr(2)) /= 77) then
    print *, "FAIL: want [77] got [", arr(2), "]"
    stop 1
end if
end program t
