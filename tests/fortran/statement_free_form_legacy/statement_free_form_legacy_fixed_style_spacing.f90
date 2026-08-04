! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_fixed_style_spacing
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

program statement_free_form_legacy_fixed_style_spacing
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ 2, 3, 4 ]
    integer :: i
    integer,parameter :: one = 1
    do i=1,3
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 3) then
            print *, "FAIL: more than 3 line(s)"
            stop 1
        end if
        if ((i + one) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", i + one, "]"
            stop 1
        end if
    end do
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program statement_free_form_legacy_fixed_style_spacing
