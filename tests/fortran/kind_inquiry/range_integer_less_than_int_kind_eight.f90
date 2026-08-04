! vybe-test: fortran/kind_inquiry/range_integer_less_than_int_kind_eight
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer :: s = 0
integer(kind=8) :: b = 0_8
if ((range(s) < range(b)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", range(s) < range(b), "]"
    stop 1
end if
end program t
