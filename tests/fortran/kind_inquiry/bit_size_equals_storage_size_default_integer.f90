! vybe-test: fortran/kind_inquiry/bit_size_equals_storage_size_default_integer
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer :: x = 0
if ((bit_size(x) == storage_size(x)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", bit_size(x) == storage_size(x), "]"
    stop 1
end if
end program t
