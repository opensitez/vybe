! vybe-test: fortran/do_concurrent_extended/do_concurrent_character_array
! origin: languages/fortran/tests/fortran/test_do_concurrent_extended.rs
program t
integer :: vybe_check_i = 0
character(len=1) :: vybe_check_w(1) = [ "b" ]
character(len=1) :: chars(3)
do concurrent (i = 1:3)
chars(i) = achar(96 + i)
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim(chars(2)) /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", chars(2), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
