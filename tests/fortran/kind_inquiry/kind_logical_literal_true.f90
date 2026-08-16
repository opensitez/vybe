! vybe-test: fortran/kind_inquiry/kind_logical_literal_true
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((kind(.true.)) /= 4) then
    print *, "FAIL: want [4] got [", kind(.true.), "]"
    stop 1
end if
end program t
