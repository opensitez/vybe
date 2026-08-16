! vybe-test: fortran/strings_extended/local_function_string_slice_runtime
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
integer :: vybe_check_i = 0
character(len=2) :: vybe_check_w(1) = [ "ab" ]
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim(trim(str_upper('ab'))) /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", trim(str_upper('ab')), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
contains
pure function str_upper(s) result(u)
character(len=*), intent(in) :: s
character(len=len(s)) :: u
integer :: i
do i = 1, len(s)
    u(i:i) = s(i:i)
end do
end function str_upper
end program t
