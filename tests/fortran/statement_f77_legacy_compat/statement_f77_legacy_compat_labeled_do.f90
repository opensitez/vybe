! vybe-test: fortran/statement_f77_legacy_compat/statement_f77_legacy_compat_labeled_do
! origin: languages/fortran/tests/fortran/test_statement_f77_legacy_compat.rs
program statement_f77_legacy_compat_labeled_do
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 6 ]
integer i
integer sum
sum = 0
do 10 i = 1, 3
sum = sum + i
10          continue
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if ((sum) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", sum, "]"
    stop 1
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program statement_f77_legacy_compat_labeled_do
