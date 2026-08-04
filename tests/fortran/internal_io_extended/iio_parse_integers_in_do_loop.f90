! vybe-test: fortran/internal_io_extended/iio_parse_integers_in_do_loop
! origin: languages/fortran/tests/fortran/test_internal_io_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
character(len=4) :: vals(3) = ['1', '2', '3']
integer :: i, n, total
 total = 0
do i = 1, 3
read(vals(i), '(I0)') n
total = total + n
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((total) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", total, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
