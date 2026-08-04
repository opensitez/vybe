! vybe-test: fortran/statement_free_form_legacy/statement_free_form_legacy_common_block_and_data
! origin: languages/fortran/tests/fortran/test_statement_free_form_legacy.rs

    program statement_free_form_legacy_common_block_and_data
    integer :: a, b, idx
    common /legacy_com/ a, b
    data a, b /1, 2/
    idx = a + b
    if ((idx) /= 3) then
    print *, "FAIL: want [3] got [", idx, "]"
    stop 1
end if
end program statement_free_form_legacy_common_block_and_data
