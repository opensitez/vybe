! vybe-test: fortran/kind_inquiry/kind_complex_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
complex :: c
c = (1.0, 2.0)
if ((kind(c)) /= 8) then
    print *, "FAIL: want [8] got [", kind(c), "]"
    stop 1
end if
end program t
