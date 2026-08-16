! vybe-test: fortran/kind_inquiry/logical_kind_and_storage_size_kind
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
logical :: l
if ((kind(l)) /= 4) then
    print *, "FAIL: want [4] got [", kind(l), "]"
    stop 1
end if
if ((kind(storage_size(l))) /= 4) then
    print *, "FAIL: want [4] got [", kind(storage_size(l)), "]"
    stop 1
end if
end program t
