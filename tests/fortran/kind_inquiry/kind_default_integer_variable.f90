! vybe-test: fortran/kind_inquiry/kind_default_integer_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer :: n = 7
if ((kind(n)) /= 8) then
    print *, "FAIL: want [8] got [", kind(n), "]"
    stop 1
end if
end program t
