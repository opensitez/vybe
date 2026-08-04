! vybe-test: fortran/kind_parameter_defaulting/test_selected_real_kind_defaulting_with_unavailable_range_is_valid
! origin: languages/fortran/tests/fortran/test_kind_parameter_defaulting.rs

program test_kind_parameter_defaulting
    if ((selected_real_kind(6)) /= 8) then
    print *, "FAIL: want [8] got [", selected_real_kind(6), "]"
    stop 1
end if
    if ((selected_real_kind(6, 38)) /= 8) then
    print *, "FAIL: want [8] got [", selected_real_kind(6, 38), "]"
    stop 1
end if
    if ((selected_real_kind(15)) /= 8) then
    print *, "FAIL: want [8] got [", selected_real_kind(15), "]"
    stop 1
end if
end program test_kind_parameter_defaulting
