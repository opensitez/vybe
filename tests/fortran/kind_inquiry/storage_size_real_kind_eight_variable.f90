! vybe-test: fortran/kind_inquiry/storage_size_real_kind_eight_variable
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
real(kind=8) :: x = 0.0_8
if ((storage_size(x)) /= 64) then
    print *, "FAIL: want [64] got [", storage_size(x), "]"
    stop 1
end if
end program t
