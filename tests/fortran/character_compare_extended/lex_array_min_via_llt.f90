! vybe-test: fortran/character_compare_extended/lex_array_min_via_llt
! origin: languages/fortran/tests/fortran/test_character_compare_extended.rs
program t
integer :: vybe_check_i = 0
character(len=3) :: vybe_check_w(1) = [ "bat" ]
character(len=3) :: a(3) = ['cat','dog','bat']
character(len=3) :: m
integer :: i
m = a(1)
do i = 2, 3
if (llt(a(i), m)) m = a(i)
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim(trim(m)) /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", trim(m), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
