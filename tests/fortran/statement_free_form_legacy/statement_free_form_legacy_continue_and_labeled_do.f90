! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_continue_and_labeled_do
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

program statement_free_form_legacy_continue_and_labeled_do
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 1, 2 ]
    integer :: i
    i = 0
    do 100 i = 1, 2
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if ((i) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", i, "]"
            stop 1
        end if
100 continue
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program statement_free_form_legacy_continue_and_labeled_do
