! vybe-test: fortran/kinds/selected_int_kind_and_kind_runtime
! origin: languages/fortran/tests/fortran/test_kinds.rs

program test
    if ((kind(1)) /= 4) then
    print *, "FAIL: want [4] got [", kind(1), "]"
    stop 1
end if
    if ((selected_int_kind(9)) /= 4) then
    print *, "FAIL: want [4] got [", selected_int_kind(9), "]"
    stop 1
end if
    if ((selected_real_kind(15, 307)) /= 8) then
    print *, "FAIL: want [8] got [", selected_real_kind(15, 307), "]"
    stop 1
end if
end program test
