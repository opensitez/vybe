! vybe-test: fortran/kind_inquiry/storage_size_real_kind_four_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real(kind=4) :: x = 0.0_4
if ((storage_size(x)) /= 32) then
    print *, "FAIL: want [32] got [", storage_size(x), "]"
    stop 1
end if
end program t
