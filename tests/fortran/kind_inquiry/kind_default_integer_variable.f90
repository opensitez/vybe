! vybe-test: fortran/kind_inquiry/kind_default_integer_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer :: n = 7
if ((kind(n)) /= 4) then
    print *, "FAIL: want [4] got [", kind(n), "]"
    stop 1
end if
end program t
