! vybe-test: fortran/do_concurrent_extended/do_concurrent_logical_all_false
! origin: languages/fortran/tests/fortran/test_do_concurrent_extended.rs
program t
integer :: vybe_check_i = 0
logical :: vybe_check_w(1) = [ .false. ]
logical :: flags(4)
do concurrent (i = 1:4)
flags(i) = .false.
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((flags(1)) .neqv. vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", flags(1), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
