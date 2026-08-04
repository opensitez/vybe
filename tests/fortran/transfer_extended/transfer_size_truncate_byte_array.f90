! vybe-test: fortran/transfer_extended/transfer_size_truncate_byte_array
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: n = 305419896
integer(kind=1) :: full(4), part(2)
full = transfer(n, full)
part = transfer(full, part, 2)
if ((int(part(1))) /= 120) then
    print *, "FAIL: want [120] got [", int(part(1)), "]"
    stop 1
end if
if ((int(part(2))) /= 86) then
    print *, "FAIL: want [86] got [", int(part(2)), "]"
    stop 1
end if
end program t
