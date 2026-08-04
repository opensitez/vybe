! vybe-test: fortran/kind_inquiry/range_int_kind_four_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer(kind=4) :: x = 0_4
if ((range(x)) /= 9) then
    print *, "FAIL: want [9] got [", range(x), "]"
    stop 1
end if
end program t
