! vybe-test: fortran/kind_inquiry/kind_int_kind_eight_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer(kind=8) :: n = 7_8
if ((kind(n)) /= 8) then
    print *, "FAIL: want [8] got [", kind(n), "]"
    stop 1
end if
end program t
