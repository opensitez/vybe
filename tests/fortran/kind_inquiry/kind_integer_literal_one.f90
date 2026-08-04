! vybe-test: fortran/kind_inquiry/kind_integer_literal_one
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((kind(1)) /= 8) then
    print *, "FAIL: want [8] got [", kind(1), "]"
    stop 1
end if
end program t
