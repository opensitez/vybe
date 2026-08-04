! vybe-test: fortran/transfer_extended/transfer_size_on_character_string
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=4) :: s = 'WXYZ'
character(len=2) :: t
t = transfer(s, t, 2)
if ((ichar(t(1:1))) /= 87) then
    print *, "FAIL: want [87] got [", ichar(t(1:1)), "]"
    stop 1
end if
if ((ichar(t(2:2))) /= 88) then
    print *, "FAIL: want [88] got [", ichar(t(2:2)), "]"
    stop 1
end if
end program t
