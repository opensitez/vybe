! vybe-test: fortran/kind_inquiry/range_int_kind_two_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer(kind=2) :: x = 0_2
if ((range(x)) /= 4) then
    print *, "FAIL: want [4] got [", range(x), "]"
    stop 1
end if
end program t
