! vybe-test: fortran/kind_inquiry/storage_size_default_real_literal
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((storage_size(0.0)) /= 32) then
    print *, "FAIL: want [32] got [", storage_size(0.0), "]"
    stop 1
end if
end program t
