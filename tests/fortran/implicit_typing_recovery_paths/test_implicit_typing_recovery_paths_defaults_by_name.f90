! vybe-test: fortran/implicit_typing_recovery_paths/test_implicit_typing_recovery_paths_defaults_by_name
! origin: languages/fortran/tests/fortran/test_implicit_typing_recovery_paths.rs

program test_implicit_typing_recovery_paths_defaults_by_name
    x = 1.0
    y = 2.0
    z = x + y
    i = 7
    if ((nint(z)) /= 3) then
    print *, "FAIL: want [3] got [", nint(z), "]"
    stop 1
end if
    if ((i) /= 7) then
    print *, "FAIL: want [7] got [", i, "]"
    stop 1
end if
end program test_implicit_typing_recovery_paths_defaults_by_name
