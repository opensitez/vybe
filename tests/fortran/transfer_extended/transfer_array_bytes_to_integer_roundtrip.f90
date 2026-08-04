! vybe-test: fortran/transfer_extended/transfer_array_bytes_to_integer_roundtrip
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: original = 305419896
integer(kind=1) :: bytes(4)
integer :: recovered
bytes = transfer(original, bytes)
recovered = transfer(bytes, 0)
if ((recovered) /= 305419896) then
    print *, "FAIL: want [305419896] got [", recovered, "]"
    stop 1
end if
end program t
