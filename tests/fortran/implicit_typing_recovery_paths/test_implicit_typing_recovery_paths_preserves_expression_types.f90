! vybe-test: fortran/implicit_typing_recovery_paths/test_implicit_typing_recovery_paths_preserves_expression_types
! origin: languages/fortran/tests/fortran/test_implicit_typing_recovery_paths.rs

program test_implicit_typing_recovery_paths
    implicit none
    integer :: whole
    real :: fractional
    whole = 7 + 3
    fractional = real(whole) / 2.0
    if ((whole) /= 10) then
    print *, "FAIL: want [10] got [", whole, "]"
    stop 1
end if
    if ((nint(fractional)) /= 5) then
    print *, "FAIL: want [5] got [", nint(fractional), "]"
    stop 1
end if
end program test_implicit_typing_recovery_paths
