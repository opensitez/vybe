! vybe-test: fortran/kind_inquiry/precision_default_real_literal
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((precision(0.0)) /= 53) then
    print *, "FAIL: want [53] got [", precision(0.0), "]"
    stop 1
end if
end program t
