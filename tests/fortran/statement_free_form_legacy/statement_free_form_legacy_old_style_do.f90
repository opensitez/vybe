! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_old_style_do
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

program statement_free_form_legacy_old_style_do
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 10 ]
    integer :: i, s
    s = 0
    do 10 i = 1, 4
        s = s + i
10  continue
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((s) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", s, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program statement_free_form_legacy_old_style_do
