! vybe-test: fortran/kind_parameter_defaulting/test_kind_parameter_defaulting_selects_default_real_kind
! origin: languages/fortran/tests/fortran/test_kind_parameter_defaulting.rs

program test_kind_parameter_defaulting
    integer :: k
    k = selected_real_kind(6)
    if ((k) /= 4) then
    print *, "FAIL: want [4] got [", k, "]"
    stop 1
end if
end program test_kind_parameter_defaulting
