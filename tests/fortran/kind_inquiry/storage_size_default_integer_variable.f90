! vybe-test: fortran/kind_inquiry/storage_size_default_integer_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer :: x = 0
if ((storage_size(x)) /= 32) then
    print *, "FAIL: want [32] got [", storage_size(x), "]"
    stop 1
end if
end program t
