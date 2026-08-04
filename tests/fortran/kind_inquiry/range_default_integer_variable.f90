! vybe-test: fortran/kind_inquiry/range_default_integer_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer :: x = 0
if ((range(x)) /= 9) then
    print *, "FAIL: want [9] got [", range(x), "]"
    stop 1
end if
end program t
