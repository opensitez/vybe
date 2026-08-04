! vybe-test: fortran/integer_mod_division/do_loop_counts_fizzbuzz_multiples_of_15
! origin: languages/fortran/tests/fortran/test_integer_mod_division.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
integer :: i, c
c = 0
do i = 1, 100
if (mod(i, 15) == 0) c = c + 1
end do
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((c) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", c, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
