! vybe-test: fortran/implicit_typing_recovery_paths/test_implicit_typing_recovery_paths_overrides_with_explicit_implicit
! origin: languages/fortran/tests/fortran/test_implicit_typing_recovery_paths.rs

program test_implicit_typing_recovery_paths_overrides
    implicit integer(a-h, o-z)
    implicit real(i-n)
    i = 3
    r = 5.0
    x = i + r
    if ((i) /= 3) then
    print *, "FAIL: want [3] got [", i, "]"
    stop 1
end if
    if ((x) /= 8) then
    print *, "FAIL: want [8] got [", x, "]"
    stop 1
end if
end program test_implicit_typing_recovery_paths_overrides
