! vybe-test: fortran/transfer_extended/transfer_char_kind1_to_integer_kind2
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer(kind=2) :: n
n = transfer('Q', 0_2)
if ((n) /= 81) then
    print *, "FAIL: want [81] got [", n, "]"
    stop 1
end if
end program t
