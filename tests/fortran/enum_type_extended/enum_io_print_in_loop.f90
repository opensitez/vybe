! vybe-test: fortran/enum_type_extended/enum_io_print_in_loop
! origin: languages/fortran/tests/fortran/test_enum_type_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 0, 1, 2 ]
enum, bind(c)
enumerator :: V0 = 0, V1 = 1, V2 = 2
end enum
integer :: vals(3) = [V0, V1, V2]
integer :: i
do i = 1, 3
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 3) then
    print *, "FAIL: more than 3 line(s)"
    stop 1
end if
if ((vals(i)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", vals(i), "]"
    stop 1
end if
end do
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program t
