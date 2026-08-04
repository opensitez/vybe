! vybe-test: fortran/random_number_extended/random_reseed_after_many_draws
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 1 ]
integer :: seed(1) = [77]
real :: r, first
integer :: i
call random_seed(put=seed)
call random_number(first)
do i = 1, 20
  call random_number(r)
end do
call random_seed(put=seed)
call random_number(r)
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((merge(1, 0, r == first)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", merge(1, 0, r == first), "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
