! vybe-test: fortran/transfer_extended/transfer_array_character_three_chars
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=1) :: c(3) = ['f', 'o', 'r']
character(len=1) :: d(3)
d = transfer(c, d)
if ((ichar(d(1))) /= 102) then
    print *, "FAIL: want [102] got [", ichar(d(1)), "]"
    stop 1
end if
if ((ichar(d(2))) /= 111) then
    print *, "FAIL: want [111] got [", ichar(d(2)), "]"
    stop 1
end if
if ((ichar(d(3))) /= 114) then
    print *, "FAIL: want [114] got [", ichar(d(3)), "]"
    stop 1
end if
end program t
